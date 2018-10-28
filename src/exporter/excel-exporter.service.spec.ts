import { Test, TestingModule } from '@nestjs/testing';
import { ExcelExporterService } from './excel-exporter.service';

describe('ExcelExporterService', () => {
  let service: ExcelExporterService;
  beforeAll(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [ExcelExporterService],
    }).compile();
    service = module.get<ExcelExporterService>(ExcelExporterService);
  });
  it('should be defined', () => {
    expect(service).toBeDefined();
  });
});
