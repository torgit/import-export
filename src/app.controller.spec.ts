import { Test, TestingModule } from '@nestjs/testing';
import { INestApplication } from '@nestjs/common';
import { AppController } from './app.controller';
import { ExporterModule } from './exporter/exporter.module';

describe('AppController', () => {
  let app: TestingModule;

  beforeAll(async () => {
    app = await Test.createTestingModule({
      controllers: [AppController],
      imports: [ExporterModule],
    }).compile();
  });

  describe('excel', () => {
    it('should return workbook', () => {
      const appController = app.get<AppController>(AppController);
      expect(appController.exportToExcel()).toBe('Hello World!');
    });
  });

//   describe('json', () => {
//     it('should return json', () => {
//       const appController = app.get<AppController>(AppController);
//       expect(appController.exportToJson()).toBe('Hello World!');
//     });
//   });
});
