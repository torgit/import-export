import { Injectable, Logger } from '@nestjs/common';
import * as XLSX from 'xlsx';
import { isNullOrUndefined } from 'util';

export const primarySheet = 'sheet1';
export enum CellType {Label, Data};

const arrayRegex = /_isArray$/;
const arrayParts = /isArray/;
const arraySizeRegex = /\[([0-9]+)\]/;

function isEmpty(obj) {
    for(var key in obj) {
        if(obj.hasOwnProperty(key))
            return false;
    }
    return true;
}

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
        const json = this.getJSON(ws, range.s.r + 1, range.s.c);
        return json;
        // return this.flattenArray(json);

    }
    
    private getJSON(ws: XLSX.WorkSheet, fromRow: number, fromCol: number, arraySize?: number, parentName?: string): any[] {
        const range = XLSX.utils.decode_range(ws['!ref']);
        const arrayName = fromCol > 0
        ? this.removeArrIndicator(ws[XLSX.utils.encode_cell({c: fromCol - 1, r: 0})].w)
        : undefined;

        var elems: any[] = [];
        for (var row = fromRow; row <= range.e.r; row++) {
            var elem = {};
            for (var col = fromCol; col <= range.e.c; col++) {
                const keyAddr = XLSX.utils.encode_cell({c: col, r: 0});
                const valueAddr = XLSX.utils.encode_cell({c: col, r: row});
                var key: string = ws[keyAddr].w;
                const value = ws[valueAddr] ? ws[valueAddr].v : undefined;
                if (value) {
                    ws[valueAddr].v = undefined;
                }
                const isArray = !isNullOrUndefined(arrayRegex.exec(key));
                const isArrayParts = !isNullOrUndefined(arrayParts.exec(key));
                const hasArraySize = !isNullOrUndefined(arraySizeRegex.exec(value));

                key = this.removeArrIndicator(key);
                // if (isArrayParts) {
                //     parentName = arrayName;
                // }
                //Array's attributes ended
                if (arrayName && !key.includes(arrayName)) {
                    col = range.e.c + 1;
                }
                const nestedKeys = key.split('.');

                const nestedObj = nestedKeys.length > 1 && !isArray && value
                    ? this.getNestedObject(nestedKeys, value) : undefined;

                if (nestedObj) {
                    elem = this.mergeNestedObjects(elem, nestedObj);
                } else if (isArray && hasArraySize) {
                    const size = +arraySizeRegex.exec(value)[1];
                    const array = this.getJSON(ws, row + 1, col + 1, size, arrayName);
                    const flattened = this.flattenArray(array);
                    const merged = this.mergeNestedObjects(elem, flattened);
                    elem = merged;
                } else {
                    const obj = {};
                    obj[key] = value;
                    elem = this.mergeNestedObjects(elem, obj)
                }
            }
            if (arraySize && arraySize === elems.length) {
                row = range.e.r + 1;
            }
            if (!isEmpty(elem)) {
                elems.push(elem);
            }
        }
        return this.flattenArray(elems);
    }

    private flattenArray(array: Array<any>) {
        if (array.length > 0) {
            const flattened = array.reduce((r1, r2) => {
                const keysR1 = Object.keys(r1);
                const keysR2 = Object.keys(r2);
                if (keysR1.length === 1 && keysR2.length === 1 && keysR1[0] === keysR2[0]) {
                    return this.mergeNestedObjects(r1, r2)
                } else {
                    if (r1 instanceof Array) {
                        return [...r1, r2];
                    }
                    return [r1, r2];
                }
            });
            return flattened;
        } else {
            return array;
        }
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
            var mainVal = mainObj[k];
            var nestedVal = nestedObj[k];
            if (nestedVal && (mainVal instanceof Array)) {
                mainObj[k] = mainVal
                ? [...mainVal, nestedVal]
                : nestedVal;
            }
            else if (nestedVal && (mainVal instanceof Object) && this.hasSameKeys(mainVal, nestedVal)) {
                mainObj[k] = mainVal
                ? [mainVal, nestedVal]
                : nestedVal;
            }
            else if (nestedVal && (mainVal instanceof Object)) {
                if(mainVal) {
                    mainObj[k] = this.mergeNestedObjects(mainVal, nestedVal)
                } else {
                    mainObj[k] = nestedVal
                }
            }
            else if (nestedVal && nestedVal instanceof Object) {
                mainObj[k] = nestedVal;
            } else if (nestedVal) {
                mainObj[k] = nestedVal;
            }
        }
        return mainObj;
    }

    private removeArrIndicator(key: string): string {
        if (key.includes('_isArray')) {
            return this.removeArrIndicator(key.replace('_isArray', ''));
        } else {
            return key
        }
    }

    private hasSameKeys(obj1: Object, obj2: Object): boolean {
        const keys1 = Object.keys(obj1);
        const keys2 = Object.keys(obj2);
        return keys1.length > 1 && keys2.length > 1 && keys1.every(k1 => keys2.indexOf(k1) > -1);
    }
}