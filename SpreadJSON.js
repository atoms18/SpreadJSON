
const XLSX = require('xlsx');
const fs = require('fs');

const ReadWriteLock = require('rwlock');

const _LOCK = new ReadWriteLock();
const writeFileLock = (path, data) => {
    _LOCK.writeLock(path, (release) => {
        fs.writeFileSync(path, data);
        release();
    });
}

class SpreadJSON {
    constructor(dbname, wb_name) {
        this.workbook_path = "./db/"+dbname+"/"+wb_name+'.json';
        this.loadWorkbook();
    }

    getSheetByName(name) {
        return new SheetJSON(this, name);
    }

    static openByName(dbname, wb_name) {
        return new SpreadJSON(dbname, wb_name);
    }

    async saveWorkbook() {
        writeFileLock(this.workbook_path, JSON.stringify(this.workbook));
    }

    loadWorkbook() {
        this.workbook = JSON.parse(fs.readFileSync(this.workbook_path));
    }

}

class SheetJSON {
    constructor(spredjson, sheetname) {
        this.spredjson = spredjson;
        this.sheetname = sheetname;
        this.worksheet = this.spredjson.workbook.Sheets[sheetname];
    }

    getRange(rowORA1, col=-1, numRow=1, numCol=1) {
        let range;
        if(col > 0) {
            range = {s: {c: col-1, r: rowORA1-1}, e: {c: col-1, r: rowORA1-1}};
        } else {
            range = XLSX.utils.decode_range(rowORA1);
        }
        range.e.r += numRow-1;
        range.e.c += numCol-1;
        return new RangeJSON(this, range);
    }

    getLastRow() {
        let range = this.worksheet['!ref'];
        if(range === undefined) range = "A1";
        
        const vals = this.getRange(range).getValues().flatMap(x => x.join(''));
        const lastRow = vals.findLastIndex((element) => element !== "") + 1;
        return lastRow;
    }

    getLastColumn() {
        let range = this.worksheet['!ref'];
        if(range === undefined) range = "A1";
        
        let vals = this.getRange(range).getValues();
        const lastCol = Math.max(...vals.map(arr => arr.findLastIndex(val => val !== "" && val !== undefined) + 1));
        return lastCol;
    }

    async saveWorksheet() {
        this.spredjson.workbook.Sheets[this.sheetname] = this.worksheet;
        await this.spredjson.saveWorkbook();
    }

    loadWorksheet() {
        this.spredjson.loadWorkbook();
        this.worksheet = this.spredjson.workbook.Sheets[this.sheetname];
    }

}

class RangeJSON {
    
    constructor(sheetjson, range) {
        this.sheetjson = sheetjson;
        this.range = range;
        
        this.row_len = this.range.e.r-this.range.s.r+1;
        this.col_len = this.range.e.c-this.range.s.c+1;
    }

    offset(row, col, numRow=0, numCol=0) {
        let new_range = JSON.parse(JSON.stringify(this.range));
        new_range.s.r += row;
        new_range.s.c += col;

        new_range.e.r = numRow <= 0 ? new_range.e.r+row : new_range.s.r+numRow-1;
        new_range.e.c = numCol <= 0 ? new_range.e.c+col : new_range.s.c+numCol-1;
        
        return new RangeJSON(this.sheetjson, new_range);
    }

    getCell(row, col) {
        const selected_row = row-1+this.range.s.r;
        const selected_col = col-1+this.range.s.c;
        if(selected_row > this.range.e.r || selected_col > this.range.e.c) throw Error("getCell() end dimensions out of bound");
        if(selected_row < this.range.s.r || selected_col < this.range.s.c) throw Error("getCell() origin dimensions out of bound");
        return new RangeJSON(this.sheetjson, {s: {c: selected_col, r: selected_row}, e: {c: selected_col, r: selected_row}});
    }

    async setValue(val) {
        const vals = Array.from(Array(this.row_len), _ => Array(this.col_len).fill(val));
        XLSX.utils.sheet_add_aoa(this.sheetjson.worksheet, vals, {origin: this.range.s});
        await this.sheetjson.saveWorksheet();
    }
    async setValues(vals) {
        if(this.row_len != vals.length || this.col_len != vals[0].length) throw Error("setValues() range dimensions != input dimensions");

        XLSX.utils.sheet_add_aoa(this.sheetjson.worksheet, vals, {origin: this.range.s});
        await this.sheetjson.saveWorksheet();
    }

    getValue() {
        this.sheetjson.loadWorksheet();
        const val = XLSX.utils.sheet_to_json(this.sheetjson.worksheet, {header: 1, range: this.range, defval: ""});
        return val[0][0];
    }
    getValues() {
        this.sheetjson.loadWorksheet();
        const val = XLSX.utils.sheet_to_json(this.sheetjson.worksheet, {header: 1, range: this.range, defval: ""});
        return val;
    }

    getLastColumn() {
        return this.range.e.c + 1;
    }

    getColumn() {
        return this.range.s.c + 1;
    }

    getNumColumns() {
        return this.col_len;
    }

    getFirstEmptyColumn() {
        const values = this.getValues()[0];
        let vals_empty_index = values.indexOf("");
        if(vals_empty_index === -1) {
            vals_empty_index = values.length;
        }
        vals_empty_index++;
        return vals_empty_index;
    }
}

exports.SpreadJSON = SpreadJSON
