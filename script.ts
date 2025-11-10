import ExcelJS from 'exceljs';
import { promises as fs } from 'fs';
import path from 'path';

type LogLevel = 'info' | 'warn' | 'error' | 'success';

interface RawRecord {
  [key: string]: ExcelJS.CellValue;
}

interface NormalizedRecord {
  id: string;
  name: string;
  email: string;
  country: string;
  age: number;
  revenue: number;
  lastPurchase: Date | null;
}

interface SummaryReport {
  totalRows: number;
  validRows: number;
  invalidRows: number;
  revenueByCountry: Record<string, number>;
  averageAge: number;
  topCustomers: NormalizedRecord[];
}

class Logger {
  constructor(private silent = false) {}

  private write(level: LogLevel, message: string) {
    if (this.silent && level !== 'error') {
      return;
    }

    const prefix = {
      info: '[ℹ️ ]',
      warn: '[⚠️ ]',
      error: '[❌]',
      success: '[✅]',
    } satisfies Record<LogLevel, string>;

    const colorCode = {
      info: '\x1b[36m',
      warn: '\x1b[33m',
      error: '\x1b[31m',
      success: '\x1b[32m',
    } satisfies Record<LogLevel, string>;

    const reset = '\x1b[0m';
    console.log(`${colorCode[level]}${prefix[level]} ${message}${reset}`);
  }

  info(message: string) {
    this.write('info', message);
  }

  warn(message: string) {
    this.write('warn', message);
  }

  error(message: string) {
    this.write('error', message);
  }

  success(message: string) {
    this.write('success', message);
  }
}

class ExcelProcessor {
  private logger: Logger;

  constructor(private readonly options: { silent?: boolean } = {}) {
    this.logger = new Logger(options.silent);
  }

  async process(filePath: string) {
    this.logger.info(`Iniciando procesamiento del archivo: ${filePath}`);

    await this.ensureFileExists(filePath);

    const workbook = await this.loadWorkbook(filePath);
    const worksheet = workbook.worksheets[0];

    if (!worksheet) {
      throw new Error('El archivo no contiene hojas de cálculo.');
    }

    this.logger.info(`Leyendo hoja: ${worksheet.name}`);

    const records = this.extractRecords(worksheet);
    const { validRecords, invalidCount } = this.normalizeRecords(records);
    const summary = this.buildSummary(validRecords, records.length, invalidCount);

    this.logger.success('Procesamiento completado correctamente.');

    return { validRecords, summary };
  }

  async exportResults(
    data: NormalizedRecord[],
    summary: SummaryReport,
    outputDir: string,
  ) {
    await fs.mkdir(outputDir, { recursive: true });

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const dataFile = path.join(outputDir, `dataset-${timestamp}.json`);
    const summaryFile = path.join(outputDir, `summary-${timestamp}.json`);

    await fs.writeFile(dataFile, JSON.stringify(data, null, 2), 'utf8');
    this.logger.success(`Datos exportados en: ${dataFile}`);

    await fs.writeFile(summaryFile, JSON.stringify(summary, null, 2), 'utf8');
    this.logger.success(`Reporte exportado en: ${summaryFile}`);
  }

  private async ensureFileExists(filePath: string) {
    try {
      const stats = await fs.stat(filePath);
      if (!stats.isFile()) {
        throw new Error(`La ruta "${filePath}" no es un archivo.`);
      }
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
        throw new Error(`No se encontró el archivo: ${filePath}`);
      }
      throw error;
    }
  }

  private async loadWorkbook(filePath: string) {
    const workbook = new ExcelJS.Workbook();
    workbook.created = new Date();
    await workbook.xlsx.readFile(filePath);
    return workbook;
  }

  private extractRecords(worksheet: ExcelJS.Worksheet): RawRecord[] {
    const headerRow = worksheet.getRow(1);
    const headers = headerRow.values
      .slice(1)
      .map(value => (typeof value === 'string' ? value.trim() : String(value)));

    if (!headers.length) {
      throw new Error('No se detectaron encabezados en la primera fila.');
    }

    const records: RawRecord[] = [];

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        return;
      }

      const rowValues = row.values.slice(1);
      const record: RawRecord = {};

      headers.forEach((header, index) => {
        record[header] = rowValues[index] ?? null;
      });

      records.push(record);
    });

    if (!records.length) {
      this.logger.warn('No se encontraron datos para procesar.');
    } else {
      this.logger.info(`Se extrajeron ${records.length} filas de datos.`);
    }

    return records;
  }

  private normalizeRecords(records: RawRecord[]) {
    const validRecords: NormalizedRecord[] = [];
    let invalidCount = 0;

    for (const record of records) {
      const normalized = this.normalizeRecord(record);
      if (normalized) {
        validRecords.push(normalized);
      } else {
        invalidCount += 1;
      }
    }

    if (invalidCount > 0) {
      this.logger.warn(`Se descartaron ${invalidCount} filas por datos inválidos.`);
    }

    return { validRecords, invalidCount };
  }

  private normalizeRecord(record: RawRecord): NormalizedRecord | null {
    const id = this.coerceString(record['id'] ?? record['ID']);
    const name = this.coerceString(record['name'] ?? record['Nombre']);
    const email = this.coerceString(record['email'] ?? record['Email']);
    const country = this.coerceString(record['country'] ?? record['País'] ?? record['Country']);
    const age = this.coerceNumber(record['age'] ?? record['Edad']);
    const revenue = this.coerceNumber(record['revenue'] ?? record['Ventas'] ?? record['Revenue']);
    const lastPurchase = this.coerceDate(record['last_purchase'] ?? record['Última compra']);

    if (!id || !name || !email || !country || age === null || revenue === null) {
      return null;
    }

    return { id, name, email, country, age, revenue, lastPurchase };
  }

  private coerceString(value: ExcelJS.CellValue): string | null {
    if (typeof value === 'string') {
      const trimmed = value.trim();
      return trimmed.length > 0 ? trimmed : null;
    }

    if (typeof value === 'number' || typeof value === 'boolean') {
      return String(value);
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    return null;
  }

  private coerceNumber(value: ExcelJS.CellValue): number | null {
    if (typeof value === 'number') {
      return Number.isFinite(value) ? value : null;
    }

    if (typeof value === 'string') {
      const parsed = Number(value.replace(/[^\d.-]/g, ''));
      return Number.isFinite(parsed) ? parsed : null;
    }

    return null;
  }

  private coerceDate(value: ExcelJS.CellValue): Date | null {
    if (value instanceof Date) {
      return value;
    }

    if (typeof value === 'number') {
      // Excel almacena fechas como números (días desde 1900-01-00)
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const millisPerDay = 24 * 60 * 60 * 1000;
      return new Date(excelEpoch.getTime() + value * millisPerDay);
    }

    if (typeof value === 'string') {
      const parsed = new Date(value);
      return Number.isNaN(parsed.getTime()) ? null : parsed;
    }

    return null;
  }

  private buildSummary(
    validRecords: NormalizedRecord[],
    totalRows: number,
    invalidRows: number,
  ): SummaryReport {
    const revenueByCountry = this.aggregateByCountry(validRecords);
    const averageAge = this.computeAverageAge(validRecords);
    const topCustomers = this.getTopCustomers(validRecords, 5);

    return {
      totalRows,
      validRows: validRecords.length,
      invalidRows,
      revenueByCountry,
      averageAge,
      topCustomers,
    };
  }

  private aggregateByCountry(records: NormalizedRecord[]) {
    return records.reduce<Record<string, number>>((acc, record) => {
      const previous = acc[record.country] ?? 0;
      acc[record.country] = previous + record.revenue;
      return acc;
    }, {});
  }

  private computeAverageAge(records: NormalizedRecord[]) {
    if (records.length === 0) {
      return 0;
    }

    const totalAge = records.reduce((acc, record) => acc + record.age, 0);
    return Number((totalAge / records.length).toFixed(2));
  }

  private getTopCustomers(records: NormalizedRecord[], limit: number) {
    return [...records]
      .sort((a, b) => b.revenue - a.revenue)
      .slice(0, limit);
  }
}

async function parseArguments() {
  const [, , ...args] = process.argv;

  const options: Record<string, string | boolean> = {};
  const positionals: string[] = [];

  for (const arg of args) {
    if (arg.startsWith('--')) {
      const [key, value] = arg.slice(2).split('=');
      options[key] = value ?? true;
    } else {
      positionals.push(arg);
    }
  }

  if (positionals.length === 0) {
    throw new Error(
      'Uso: ts-node script.ts <ruta-al-archivo.xlsx> [--output=carpeta] [--silent]',
    );
  }

  const filePath = positionals[0];
  const outputDir =
    typeof options.output === 'string' && options.output.length > 0
      ? options.output
      : path.dirname(path.resolve(filePath));
  const silent = Boolean(options.silent);

  return { filePath, outputDir, silent };
}

async function main() {
  try {
    const { filePath, outputDir, silent } = await parseArguments();
    const processor = new ExcelProcessor({ silent });

    const { validRecords, summary } = await processor.process(filePath);

    if (validRecords.length === 0) {
      processor['logger'].warn('No hay registros válidos para exportar.');
      return;
    }

    await processor.exportResults(validRecords, summary, outputDir);

    processor['logger'].info('Previsualización de resumen:');
    processor['logger'].info(JSON.stringify(summary, null, 2));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('\x1b[31m[❌] Error fatal:\x1b[0m', message);
    process.exitCode = 1;
  }
}

if (require.main === module) {
  void main();
}

