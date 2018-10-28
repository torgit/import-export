import { Module } from '@nestjs/common';
import { ExporterModule } from './exporter/exporter.module';
import { AppController } from './app.controller';
import { UtilsModule } from './utils/utils.module';

@Module({
  imports: [ExporterModule, UtilsModule],
  controllers: [AppController],
})
export class AppModule {}
