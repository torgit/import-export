import { Module } from '@nestjs/common';
import { ExcelExporterService } from './excel-exporter.service';
import { JsonExporterService } from './json-exporter.service';
import { UtilsModule } from '../utils/utils.module';

const ExcelExporterServiceProvider = {provide: 'ExcelExporterService', useValue: ExcelExporterService}
const JsonExporterServiceProvider = {provide: 'JsonExporterService', useValue: JsonExporterService}

@Module({
  providers: [
    ExcelExporterServiceProvider, JsonExporterServiceProvider,
    ExcelExporterService, JsonExporterService
  ],
  imports: [
    UtilsModule,
  ],
  exports: [
    ExcelExporterService, JsonExporterService
  ],
})
export class ExporterModule {}
