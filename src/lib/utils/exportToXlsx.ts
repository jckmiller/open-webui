/**
 * Client-side spreadsheet export helpers.
 *
 * Three entry points:
 *  1. exportDelimitedToXlsx  – CSV / TSV code-block → .xlsx
 *  2. exportMarkdownTableToXlsx – Markdown table token → .xlsx
 *  3. exportXlsxSpec         – ```xlsx JSON spec block → multi-sheet .xlsx (Option B)
 *
 * All write operations are purely client-side via SheetJS
 * ("xlsx": "^0.18.5" is already in package.json).
 */

import fileSaver from 'file-saver';
const { saveAs } = fileSaver;

// ─── internal helper ────────────────────────────────────────────────────────

function blobSave(data: ArrayBuffer, filename: string) {
	const name = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`;
	saveAs(new Blob([data], { type: 'application/octet-stream' }), name);
}

// ─── 1. CSV / TSV code block ─────────────────────────────────────────────────

/**
 * Export raw CSV or TSV text from a fenced code block as an .xlsx file.
 *
 * @param text      - the raw code block content
 * @param filename  - suggested download name (extension added automatically)
 * @param delimiter - ',' for CSV (default) | '\t' for TSV
 */
export async function exportDelimitedToXlsx(
	text: string,
	filename = 'export',
	delimiter = ','
): Promise<void> {
	const XLSX = await import('xlsx');
	// SheetJS can natively parse CSV/TSV via read() when type = 'string'
	const wb = XLSX.read(text, { type: 'string', FS: delimiter });
	const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
	blobSave(out, filename);
}

// ─── 2. Markdown table ───────────────────────────────────────────────────────

/**
 * Export a parsed Markdown table (header row + data rows) as an .xlsx file.
 *
 * @param headers  - array of plain-text header strings
 * @param rows     - 2-D array of plain-text cell strings
 * @param filename - suggested download name
 */
export async function exportMarkdownTableToXlsx(
	headers: string[],
	rows: string[][],
	filename = 'table'
): Promise<void> {
	const XLSX = await import('xlsx');
	const aoa = [headers, ...rows];
	const ws = XLSX.utils.aoa_to_sheet(aoa);
	const wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
	const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
	blobSave(out, filename);
}

// ─── 3. ```xlsx JSON spec (Option B) ─────────────────────────────────────────

/**
 * Shape of the JSON spec a model emits inside a ```xlsx … ``` block.
 *
 * Example the model should produce:
 *
 *   ```xlsx
 *   {
 *     "filename": "budget.xlsx",
 *     "sheets": [
 *       {
 *         "name": "Q1 Sales",
 *         "headers": ["Product", "Units", "Revenue"],
 *         "rows": [
 *           ["Widget A", 120, 1200.00],
 *           ["Widget B", 85,  892.50]
 *         ]
 *       }
 *     ]
 *   }
 *   ```
 *
 * Fields:
 *  - `filename`        (optional) – download file name; defaults to "export.xlsx"
 *  - `sheets`          (required) – one or more sheet definitions
 *  - `sheets[].name`   (optional) – sheet tab name; defaults to "Sheet1", "Sheet2", …
 *  - `sheets[].headers`(optional) – first row of column labels
 *  - `sheets[].rows`   (required) – 2-D array of cell values (string | number | boolean | null)
 */
export interface XlsxSpec {
	filename?: string;
	sheets: Array<{
		name?: string;
		headers?: (string | number | null)[];
		rows: (string | number | boolean | null)[][];
	}>;
}

/**
 * Parse a JSON ```xlsx spec and write a real multi-sheet .xlsx file.
 * Throws if the JSON is invalid or the spec shape is wrong.
 */
export async function exportXlsxSpec(specText: string): Promise<void> {
	const spec: XlsxSpec = JSON.parse(specText);

	if (!Array.isArray(spec.sheets) || spec.sheets.length === 0) {
		throw new Error('xlsx spec must have at least one sheet in "sheets" array');
	}

	const XLSX = await import('xlsx');
	const wb = XLSX.utils.book_new();

	for (const [i, sheet] of spec.sheets.entries()) {
		const aoa: (string | number | boolean | null)[][] = sheet.headers
			? [sheet.headers, ...sheet.rows]
			: sheet.rows;
		const ws = XLSX.utils.aoa_to_sheet(aoa);
		XLSX.utils.book_append_sheet(wb, ws, sheet.name ?? `Sheet${i + 1}`);
	}

	const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
	blobSave(out, spec.filename ?? 'export');
}
