import * as XLSX from 'xlsx';

const FILE_KEY = 'spreadsheet.xlsx';

// ── Tool schemas for OpenAI Realtime API ──────────────────────────────────────
export const TOOLS = [
  {
    type: 'function' as const,
    name: 'get_sheet_info',
    description: 'Get the spreadsheet file name, sheet names, and row/column counts. Call this first.',
    parameters: { type: 'object' as const, properties: {} },
  },
  {
    type: 'function' as const,
    name: 'read_cell',
    description: 'Read the value of a single cell.',
    parameters: {
      type: 'object' as const,
      properties: {
        cell: { type: 'string', description: 'Cell address like A1 or B3' },
        sheet_name: { type: 'string', description: 'Sheet tab name (optional)' },
      },
      required: ['cell'],
    },
  },
  {
    type: 'function' as const,
    name: 'write_cell',
    description: 'Write a value to a single cell. Formulas starting with = are supported.',
    parameters: {
      type: 'object' as const,
      properties: {
        cell: { type: 'string', description: 'Cell address like A1' },
        value: { type: 'string', description: 'Value or formula like =SUM(A1:A5)' },
        sheet_name: { type: 'string', description: 'Sheet tab name (optional)' },
      },
      required: ['cell', 'value'],
    },
  },
  {
    type: 'function' as const,
    name: 'read_range',
    description: 'Read all values in a cell range. Returns a 2D array.',
    parameters: {
      type: 'object' as const,
      properties: {
        range: { type: 'string', description: 'Range like A1:D10' },
        sheet_name: { type: 'string', description: 'Sheet tab name (optional)' },
      },
      required: ['range'],
    },
  },
  {
    type: 'function' as const,
    name: 'write_range',
    description: 'Write a 2D array of values starting at a given cell.',
    parameters: {
      type: 'object' as const,
      properties: {
        start_cell: { type: 'string', description: 'Top-left starting cell like A1' },
        values: {
          type: 'array',
          items: { type: 'array', items: { type: 'string' } },
          description: 'Array of rows; each row is an array of cell values',
        },
        sheet_name: { type: 'string', description: 'Sheet tab name (optional)' },
      },
      required: ['start_cell', 'values'],
    },
  },
  {
    type: 'function' as const,
    name: 'add_row',
    description: 'Append a new row at the bottom of the data.',
    parameters: {
      type: 'object' as const,
      properties: {
        values: {
          type: 'array',
          items: { type: 'string' },
          description: 'Cell values for the new row, left to right',
        },
        sheet_name: { type: 'string', description: 'Sheet tab name (optional)' },
      },
      required: ['values'],
    },
  },
  {
    type: 'function' as const,
    name: 'clear_range',
    description: 'Clear all values in a cell range.',
    parameters: {
      type: 'object' as const,
      properties: {
        range: { type: 'string', description: 'Range to clear like A1:D10' },
        sheet_name: { type: 'string', description: 'Sheet tab name (optional)' },
      },
      required: ['range'],
    },
  },
];

// ── R2 helpers ────────────────────────────────────────────────────────────────
async function loadWb(r2: R2Bucket): Promise<XLSX.WorkBook> {
  const obj = await r2.get(FILE_KEY);
  if (!obj) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), 'Sheet1');
    return wb;
  }
  const buf = await obj.arrayBuffer();
  return XLSX.read(new Uint8Array(buf), { type: 'array' });
}

async function saveWb(wb: XLSX.WorkBook, r2: R2Bucket): Promise<void> {
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as number[];
  await r2.put(FILE_KEY, new Uint8Array(buf));
}

function getSheet(wb: XLSX.WorkBook, sheetName?: string): XLSX.WorkSheet {
  const name = sheetName && wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames[0];
  return wb.Sheets[name];
}

function coerce(v: string): string | number {
  const n = Number(v);
  return isNaN(n) || v.trim() === '' ? v : n;
}

function cellObj(val: string): XLSX.CellObject {
  if (val.startsWith('=')) return { f: val.slice(1), t: 'f' };
  const c = coerce(val);
  return typeof c === 'number' ? { v: c, t: 'n' } : { v: c, t: 's' };
}

// ── Tool execution ────────────────────────────────────────────────────────────
export async function executeTool(
  name: string,
  args: Record<string, unknown>,
  r2: R2Bucket,
): Promise<unknown> {
  const wb = await loadWb(r2);

  switch (name) {
    case 'get_sheet_info': {
      return {
        file: FILE_KEY,
        sheets: wb.SheetNames.map(n => {
          const ref = wb.Sheets[n]['!ref'] ? XLSX.utils.decode_range(wb.Sheets[n]['!ref']!) : null;
          return { name: n, rows: ref ? ref.e.r - ref.s.r + 1 : 0, cols: ref ? ref.e.c - ref.s.c + 1 : 0 };
        }),
      };
    }

    case 'read_cell': {
      const ws = getSheet(wb, args.sheet_name as string);
      const addr = (args.cell as string).toUpperCase();
      const c = ws[addr] as XLSX.CellObject | undefined;
      return { cell: addr, value: c ? (c.w ?? c.v ?? null) : null };
    }

    case 'write_cell': {
      const ws = getSheet(wb, args.sheet_name as string);
      const addr = (args.cell as string).toUpperCase();
      ws[addr] = cellObj(args.value as string);
      const d = XLSX.utils.decode_cell(addr);
      const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
      ref.e.r = Math.max(ref.e.r, d.r);
      ref.e.c = Math.max(ref.e.c, d.c);
      ws['!ref'] = XLSX.utils.encode_range(ref);
      await saveWb(wb, r2);
      return { success: true, cell: addr, value: args.value };
    }

    case 'read_range': {
      const ws = getSheet(wb, args.sheet_name as string);
      const values = XLSX.utils.sheet_to_json<string[]>(ws, {
        header: 1,
        range: (args.range as string).toUpperCase(),
        defval: '',
      });
      return { range: args.range, rows: values.length, values };
    }

    case 'write_range': {
      const ws = getSheet(wb, args.sheet_name as string);
      const origin = XLSX.utils.decode_cell((args.start_cell as string).toUpperCase());
      const values = args.values as string[][];
      values.forEach((row, ri) =>
        row.forEach((val, ci) => {
          ws[XLSX.utils.encode_cell({ r: origin.r + ri, c: origin.c + ci })] = cellObj(val);
        }),
      );
      const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
      ref.e.r = Math.max(ref.e.r, origin.r + values.length - 1);
      ref.e.c = Math.max(ref.e.c, origin.c + Math.max(...values.map(r => r.length)) - 1);
      ws['!ref'] = XLSX.utils.encode_range(ref);
      await saveWb(wb, r2);
      return { success: true, startCell: (args.start_cell as string).toUpperCase(), rowsWritten: values.length };
    }

    case 'add_row': {
      const ws = getSheet(wb, args.sheet_name as string);
      const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: -1, c: 0 } };
      const newRow = ref.e.r + 1;
      (args.values as string[]).forEach((val, ci) => {
        ws[XLSX.utils.encode_cell({ r: newRow, c: ci })] = cellObj(val);
      });
      ref.e.r = newRow;
      ref.e.c = Math.max(ref.e.c, (args.values as string[]).length - 1);
      ws['!ref'] = XLSX.utils.encode_range(ref);
      await saveWb(wb, r2);
      return { success: true, rowAdded: args.values, atRow: newRow + 1 };
    }

    case 'clear_range': {
      const ws = getSheet(wb, args.sheet_name as string);
      const range = XLSX.utils.decode_range((args.range as string).toUpperCase());
      for (let r = range.s.r; r <= range.e.r; r++)
        for (let c = range.s.c; c <= range.e.c; c++)
          delete ws[XLSX.utils.encode_cell({ r, c })];
      await saveWb(wb, r2);
      return { success: true, cleared: args.range };
    }

    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}
