import { Test, TestingModule } from '@nestjs/testing';
import { JsonExporterService } from './json-exporter.service';

describe('JsonExporterService', () => {
  let service: JsonExporterService;
  beforeAll(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [JsonExporterService],
    }).compile();
    service = module.get<JsonExporterService>(JsonExporterService);
  });
  it('should be defined', () => {
    expect(service).toBeDefined();
  });
});
