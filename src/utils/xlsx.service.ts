import { Injectable } from '@nestjs/common';
import * as XLSX from 'xlsx';

export const primarySheet = 'sheet1';
export enum CellType {Label, Data};

@Injectable()
export class XlsxService {
    getNewWorkbook(): XLSX.WorkBook {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet([]);
        wb.SheetNames.push(primarySheet);
        wb.Sheets[primarySheet] = ws;
        return wb;
    }

    writeWorkbook(wb: XLSX.WorkBook, entries: number): void {
        const ws = wb.Sheets[primarySheet];
        ws['!ref'] = `A0:${ws['!ref'].charAt(0)}${entries}`;
        XLSX.writeFile(wb, 'result.xlsx');
    }

    addCellToSheet(
        ws: XLSX.WorkSheet, 
        address: string, 
        value: string | boolean | number | Date,
        type: CellType = CellType.Data
    ): void {
        const cell = {t:'?', v:value};

        /* assign type */
        if(typeof value == "string") cell.t = 's'; // string
        else if(typeof value == "number") cell.t = 'n'; // number
        else if(value === true || value === false) cell.t = 'b'; // boolean
        else if(value instanceof Date) cell.t = 'd';
        else throw new Error("cannot store value");

        /* add to worksheet, overwriting a cell if it exists */
        ws[address] = cell;
    
        /* find the cell range */
        const range = XLSX.utils.decode_range(ws['!ref']);
        const addr = XLSX.utils.decode_cell(address);
    
        /* extend the range to include the new cell */
        if(range.s.c > addr.c) range.s.c = addr.c;
        if(range.s.r > addr.r) range.s.r = addr.r;
        if(range.e.c < addr.c) range.e.c = addr.c;
        if(range.e.r < addr.r) range.e.r = addr.r;
    
        /* update range */
        if (type === CellType.Label) {
            ws['!ref'] = XLSX.utils.encode_range(range);
        }
    }

    addNewLabel(
        ws: XLSX.WorkSheet,
        labelName: string,
    ): void {
        const latestLabelAddr = ws['!ref'];
        const labelDecodedAddr = XLSX.utils.decode_cell(latestLabelAddr);
        if (labelDecodedAddr.c !== 0 || ws[latestLabelAddr]) {
            labelDecodedAddr.c += 1;
        }
        const labelAddress = XLSX.utils.encode_cell(labelDecodedAddr);
        this.addCellToSheet(ws, labelAddress, labelName, CellType.Label);
    }

    addNewEntry(
        ws: XLSX.WorkSheet,
        label: string,
        row: number,
        value: string | boolean | number | Date,
    ): void {
        const cellAddr = label.charAt(0) + row;
        this.addCellToSheet(ws, cellAddr, value);
    }

    readFile(buffer: Buffer): Object {
        const data = XLSX.read(buffer);
        const firstSheetName = data.SheetNames[0];
        const ws = data.Sheets[firstSheetName];
        const range = XLSX.utils.decode_range(ws['!ref']);
        var results = [];
        for (var r = range.s.r + 1; r <= range.e.r; r++) {
            var result = {};
            var objectSize = 0;
            var objectKeys = [];
            for (var c = range.s.c; c <= range.e.c; c++) {
                const obj = {};
                const keyAddr = XLSX.utils.encode_cell({c, r: 0});
                const valueAddr = XLSX.utils.encode_cell({c, r});
                const key: string = ws[keyAddr].w;
                const nestedKeys = key.split('.')
                const value = ws[valueAddr] ? ws[valueAddr].v : undefined;
                const nestedObj = nestedKeys.length > 1 ? this.getNestedObject(nestedKeys, value) : undefined;

                if (value) {
                    objectSize++;
                }
                if (value && nestedObj) {
                    objectKeys = [...objectKeys, ...nestedKeys];
                }

                if (nestedObj) {
                    result = this.mergeNestedObjects(result, nestedObj);
                } else {
                    obj[key] = value;
                    result = {...result, ...obj};
                    Object.assign(result, obj);
                }
            }

            // Check lone attribute (nested attribute)
            if (objectSize === 1) {
                const latestResult = results.pop();
                result = this.mergeNestedArrays(latestResult, result);
            }
            results.push(result);
        }
        return results.reduce((r1, r2) => {
            const keyR1 = Object.keys(r1);
            const keyR2 = Object.keys(r1);
            if (keyR1.length === 1 && keyR2.length === 1 && keyR1[0] === keyR2[0]) {
                return this.mergeNestedArrays(r1, r2)
            } else {
                if (r1 instanceof Array) {
                    return [...r1, r2];
                }
                return [r1, r2];
            }
        });
    }

    private getNestedObject(nestedKeys: string[], value: any, obj: Object = {}): Object {
        const [head, ...tail] = nestedKeys;
        if (tail.length > 0) {
            obj[head] = this.getNestedObject(tail, value);
            return obj;
        }
        else {
            obj[head] = value;
            return obj;
        } 
    }

    private mergeNestedObjects(mainObj: Object, nestedObj: Object): Object {
        for (var k in nestedObj) {
            if (nestedObj.hasOwnProperty(k) && nestedObj[k] && (nestedObj[k] instanceof Object)) {
                mainObj[k] = mainObj[k]
                ? this.mergeNestedObjects(mainObj[k], nestedObj[k])
                : nestedObj[k];
            }
            else if (nestedObj.hasOwnProperty(k) && nestedObj[k]) {
                const mainVal = mainObj[k];
                const mergedVal = mainVal ? [mainVal, nestedObj[k]] : nestedObj[k];
                mainObj[k] = mergedVal;
            }
        }
        return mainObj;
    }

    private mergeNestedArrays(mainObj: Object, nestedObj: Object): Object {
        for (var k in nestedObj) {
            if (nestedObj.hasOwnProperty(k) && nestedObj[k] && (nestedObj[k] instanceof Object)) {
                mainObj[k] = mainObj[k]
                ? mainObj[k] instanceof Array
                    ? [...mainObj[k], nestedObj[k]]
                    : [mainObj[k], nestedObj[k]]
                : nestedObj[k];
            }
            else if (nestedObj.hasOwnProperty(k) && nestedObj[k]) {
                const mainVal = mainObj[k];
                const mergedVal = mainVal ? [mainVal, nestedObj[k]] : nestedObj[k];
                mainObj[k] = mergedVal;
            }
        }
        return mainObj;
    }
}