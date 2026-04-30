import 'dotenv/config';
import { getGraphClient } from '../src/graph/client';

async function inspect(path: string, table: string) {
  const c = getGraphClient();
  try {
    const f: any = await c.api(`/me/drive/root:${path}`).get();
    console.log(`\n${path} — size ${f.size} bytes, modified ${f.lastModifiedDateTime}`);
    const tables: any = await c.api(`/me/drive/items/${f.id}/workbook/tables`).get();
    console.log(`  tables: ${tables.value.map((t: any) => t.name).join(', ') || '(none)'}`);
    if (tables.value.find((t: any) => t.name === table)) {
      const rows: any = await c.api(`/me/drive/items/${f.id}/workbook/tables/${table}/rows`).get();
      console.log(`  ${table} rows: ${rows.value.length}`);
      for (const r of rows.value.slice(-5)) {
        console.log('   ', JSON.stringify(r.values[0]).slice(0, 250));
      }
    }
  } catch (e: any) {
    console.log(`${path}: ${e.statusCode || ''} ${e.code || ''} ${(e.message || '').slice(0, 200)}`);
  }
}

(async () => {
  await inspect('/Documents/BillHunter/bills.xlsx', 'Bills');
  await inspect('/Documents/BillHunter/taxes_2026.xlsx', 'TaxLog');
})();
