import { Injectable } from '@nestjs/common';
import { IExporter } from './interfaces/exporter-service.interface';
import { XlsxService } from 'utils/xlsx.service';

@Injectable()
export class JsonExporterService implements IExporter<Buffer, Object> {
    constructor(
        private readonly xlsxService: XlsxService
    ) {}

    async export(file: Buffer): Promise<Object> {
        return this.xlsxService.readFile(file);
    }
}
