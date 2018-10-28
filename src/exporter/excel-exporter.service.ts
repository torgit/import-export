import { Injectable } from '@nestjs/common';
import { IExporter } from './interfaces/exporter-service.interface';
import { XlsxService, primarySheet } from '../utils/xlsx.service';
import * as XLSX from 'xlsx';

interface LabelMap {
    [name: string]: string;
}

@Injectable()
export class ExcelExporterService implements IExporter<Object, XLSX.WorkBook> {
    constructor(
        private readonly xlsxService: XlsxService
    ) {}

    async export(obj: Object): Promise<XLSX.WorkBook> {
        const wb = this.xlsxService.getNewWorkbook();
        const ws = wb.Sheets[primarySheet];
        const arr = obj instanceof Array ? obj as Array<Object> : undefined;
        const labels: LabelMap = {};
        const startRow = 2;
        const entries = arr
        ? this.addArrayToWorksheet(arr, ws, labels, startRow)
        : this.addPropertiesToWorksheet(obj, ws, labels, startRow);
        this.xlsxService.writeWorkbook(wb, entries);
        return wb;
    }

    private addArrayToWorksheet(arr: Array<Object>, ws: XLSX.WorkSheet, labels: LabelMap, row: number, parentName?: string): number {
        return arr.map(a => {
            const entries = this.addPropertiesToWorksheet(a, ws, labels, row, parentName);
            row = entries;
            return row;
        }).reduce((a, b) => a + b);
    }

    private addPropertiesToWorksheet (obj: Object, ws: XLSX.WorkSheet, labels: LabelMap, row: number, parentName?: string, latestLabelAddr: string = 'A1'): number {
        for (var key in obj) {
            const labelName = parentName ? `${parentName}.${key}` : key;

            if (obj.hasOwnProperty(key) && !(obj[key] instanceof Array)) {
                const value = obj[key];
                if (!labels[labelName]) {
                    this.xlsxService.addNewLabel(ws, labelName);
                    latestLabelAddr = ws["!ref"];
                    labels[labelName] = latestLabelAddr;
                }
                const label = labels[labelName];

                switch(typeof value) {
                    case "object": {
                        this.addPropertiesToWorksheet(value, ws, labels, row, labelName, latestLabelAddr);
                        break;
                    }
                    default: {
                        this.xlsxService.addNewEntry(ws, label, row, value);
                        break;
                    }
                }
            } else if (obj.hasOwnProperty(key)) {
                const array: Array<Object> = obj[key];
                array.map(a => {
                    this.addPropertiesToWorksheet(a, ws, labels, row, labelName, latestLabelAddr);
                    row += 1;
                });
                row -= 1;
            }
        }
        row += 1;
        return row;
    }
}
