/**
 * Regenerate the EMPTY_XLSX_BASE64 constant.
 *
 * Requires Python with openpyxl installed:
 *   pip install openpyxl
 *
 * Run: npm run build:empty-xlsx
 */

import { execSync } from 'child_process';
import { writeFileSync } from 'fs';
import path from 'path';

const py = `
import openpyxl, base64, io
wb = openpyxl.Workbook()
wb.active.title = 'Sheet1'
buf = io.BytesIO()
wb.save(buf)
print(base64.b64encode(buf.getvalue()).decode('ascii'))
`;

const b64 = execSync(`python3 -c "${py.replace(/"/g, '\\"')}"`).toString().trim();

const out = path.join(process.cwd(), 'src/lib/empty-xlsx.ts');
writeFileSync(
  out,
  `/**
 * Base64-encoded empty .xlsx workbook (Sheet1 only).
 * Used by excel.ts to bootstrap files that don't exist yet.
 * Regenerate with: npm run build:empty-xlsx
 */
export const EMPTY_XLSX_BASE64 = '${b64}';
`,
);
console.log(`Wrote ${out} (${b64.length} chars)`);
