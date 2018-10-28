import { Controller, Post, Inject, Body, FileInterceptor, UseInterceptors, UploadedFile } from '@nestjs/common';
import { IExporter } from './exporter/interfaces/exporter-service.interface';
import * as XLSX from 'xlsx';

@Controller('export')
export class AppController {
  constructor(
      @Inject('ExcelExporterService') private readonly excelExporterService: IExporter<Object, XLSX.WorkBook>, 
      @Inject('JsonExporterService') private readonly jsonExporterService: IExporter<Buffer, Object>
    ) {}

  @Post('excel')
  async exportToExcel(@Body() body: Object): Promise<string> {
    await this.excelExporterService.export(body);
    return 'Successfully export to xlsx';
  }

  @Post('json')
  @UseInterceptors(FileInterceptor('file'))
  async exportToJson(@UploadedFile() file): Promise<Object> {
    return await this.jsonExporterService.export(file.buffer);
  }
}
