import { useState, type ChangeEvent } from 'react'
import './App.css'
import parWorksheetTemplate from '../templates/par-worksheet.csv?raw'
import prepListTemplate from '../templates/production-prep-list.csv?raw'
import parWorksheetXlsxUrl from '../templates/par-worksheet.xlsx?url'
import prepListXlsxUrl from '../templates/production-prep-list.xlsx?url'
import Papa from 'papaparse'

const MVT_HEADERS: Record<number, string> = {
  0: 'Code',
  5: 'Monday',
  6: 'Tuesday',
  7: 'Wednesday',
  8: 'Thursday',
  9: 'Friday',
  10: 'Saturday',
  11: 'Sunday',
  12: 'Week mvt',
};

const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// A value written into a template, by zero-based row/column. The .csv and
// .xlsx templates share the same layout, so one edit applies to both.
type CellEdit = { row: number; col: number; value: number };

type OutputFile = {
  label: string;
  fileName: string;
  csvUrl: string;
  xlsxUrl: string;
};

function parseCsv(content: string): string[][] {
  return Papa.parse<string[]>(content.trim()).data;
}

function hasMovementHeaders(content: string): boolean {
  const headers = Papa.parse<string[]>(content.trim(), { preview: 1 }).data[0] ?? [];
  return Object.entries(MVT_HEADERS).every(
    ([index, name]) => headers[Number(index)]?.trim().toLowerCase() === name.toLowerCase()
  );
}

// Spreadsheet-style column name for a zero-based index (0 -> A, 12 -> M)
function columnLetter(index: number): string {
  return String.fromCharCode(65 + index);
}

function applyEdits(rows: string[][], edits: CellEdit[]) {
  for (const { row, col, value } of edits) rows[row][col] = String(value);
}

function buildCsv(rows: string[][]): Blob {
  return new Blob([Papa.unparse(rows)], { type: 'text/csv' });
}

async function buildXlsx(templateUrl: string, edits: CellEdit[]): Promise<Blob> {
  // loaded on demand: ExcelJS is large and only needed once pars are generated
  const { default: ExcelJS } = await import('exceljs');
  const workbook = new ExcelJS.Workbook();
  const template = await fetch(templateUrl).then(res => res.arrayBuffer());
  await workbook.xlsx.load(template);

  const sheet = workbook.worksheets[0];
  for (const { row, col, value } of edits) {
    sheet.getCell(row + 1, col + 1).value = value;
  }

  return new Blob([await workbook.xlsx.writeBuffer()], { type: XLSX_TYPE });
}

function revokeOutputs(outputs: OutputFile[] | null) {
  for (const output of outputs ?? []) {
    URL.revokeObjectURL(output.csvUrl);
    URL.revokeObjectURL(output.xlsxUrl);
  }
}

function App() {
  const [movementReportContent, setMovementReport] = useState('');
  const [errorStatus, setError] = useState(0);
  const [headerWarning, setHeaderWarning] = useState(false);
  const [showHeaderInfo, setShowHeaderInfo] = useState(false);
  const [outputs, setOutputs] = useState<OutputFile[] | null>(null);

  function readMovementReport(event: ChangeEvent<HTMLInputElement>) {
    revokeOutputs(outputs);
    setOutputs(null);

    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = function () {
        const content = reader.result as string;
        setMovementReport(content);
        setHeaderWarning(!hasMovementHeaders(content));
        setShowHeaderInfo(false);
      };
      reader.readAsText(file);
    }
  }

  async function generatePars() {
    if (movementReportContent === '') {
      setError(1);
      return;
    }

    setError(0);

    const mvtRows = parseCsv(movementReportContent);
    const pwRows = parseCsv(parWorksheetTemplate);

    const weekQtyByCode = new Map<string, number>();
    for (const row of mvtRows.slice(1)) {
      const code = row[0]?.trim();
      if (!code) continue;

      const weekMvt = Number(row[12]);
      const avgPerDay = weekMvt / 7;
      const zeroDays = row.slice(5, 12).filter(v => Number(v) === 0).length;

      // adjust for days with zero movement
      const weekQty = weekMvt + zeroDays * avgPerDay;
      weekQtyByCode.set(code, Math.round(weekQty));
    }

    const pwEdits: CellEdit[] = [];
    pwRows.forEach((row, i) => {
      if (i < 1) return;
      const code = row[0]?.trim();
      if (code && weekQtyByCode.has(code)) {
        const weekQty = weekQtyByCode.get(code)!;
        pwEdits.push({ row: i, col: 4, value: weekQty });
        const avgPerDay = weekQty / 7;
        pwEdits.push({ row: i, col: 6, value: avgPerDay });

        const shelfLife = row[2]?.trim();
        if (shelfLife) pwEdits.push({ row: i, col: 7, value: Math.round(avgPerDay * Number(shelfLife)) });
      }
    });
    applyEdits(pwRows, pwEdits);

    const approxParByCode = new Map<string, string>();
    for (const row of pwRows.slice(1)) {
      const code = row[0]?.trim();
      if (code) approxParByCode.set(code, row[7]);
    }

    const plRows = parseCsv(prepListTemplate);
    const plEdits: CellEdit[] = [];
    plRows.forEach((row, i) => {
      if (i < 3) return;
      for (const [pluIndex, parIndex] of [[1, 3], [9, 11]] as const) {
        const code = row[pluIndex]?.trim();
        if (code && approxParByCode.has(code)) {
          const approxPar = approxParByCode.get(code)!;
          if (approxPar !== '') plEdits.push({ row: i, col: parIndex, value: Number(approxPar) });
        } else if (code && weekQtyByCode.has(code)) {
          // not on the par worksheet: using a one-week par in absence of shelf life value
          plEdits.push({ row: i, col: parIndex, value: weekQtyByCode.get(code)! });
        }
      }
    });
    applyEdits(plRows, plEdits);

    try {
      const [pwXlsx, plXlsx] = await Promise.all([
        buildXlsx(parWorksheetXlsxUrl, pwEdits),
        buildXlsx(prepListXlsxUrl, plEdits),
      ]);

      revokeOutputs(outputs);
      setOutputs([
        {
          label: 'Par Worksheet',
          fileName: 'par-worksheet',
          csvUrl: URL.createObjectURL(buildCsv(pwRows)),
          xlsxUrl: URL.createObjectURL(pwXlsx),
        },
        {
          label: 'Production Prep List',
          fileName: 'production-prep-list',
          csvUrl: URL.createObjectURL(buildCsv(plRows)),
          xlsxUrl: URL.createObjectURL(plXlsx),
        },
      ]);
    } catch (err) {
      console.error(err);
      setError(2);
    }
  }

  return (
    <>
      <header className="app-header">
        <h1>Production Par</h1>
      </header>
      <main className="card">
        <section className="step">
          <h2>1. Upload Movement Report</h2>
          <label htmlFor="fileInput">
            <input type="file" id="fileInput" accept=".csv" onChange={readMovementReport} />
          </label>
          {headerWarning && (
            <div className="warning">
              <p>This document doesn't appear to have the correct headers. Outputs may not generate correctly</p>
              <button
                type="button"
                className="link-button"
                aria-expanded={showHeaderInfo}
                aria-controls="header-info"
                onClick={() => setShowHeaderInfo(show => !show)}
              >
                {showHeaderInfo ? 'Less info' : 'More info'}
              </button>
              {showHeaderInfo && (
                <div id="header-info" className="header-info">
                  <p>The first row of the movement report should contain these headers:</p>
                  <table>
                    <thead>
                      <tr>
                        <th>Column</th>
                        <th>Header</th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.entries(MVT_HEADERS).map(([index, name]) => (
                        <tr key={index}>
                          <td>{columnLetter(Number(index))}</td>
                          <td>{name}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <p>Other columns can contain anything. Header names aren't case-sensitive.</p>
                </div>
              )}
            </div>
          )}
        </section>

        <section className="step">
          <h2>2. Generate</h2>
          <button className="button" onClick={generatePars}>Generate Pars</button>
          {errorStatus === 1 && <p className="error">No file inputted</p>}
          {errorStatus === 2 && <p className="error">Something went wrong generating the files</p>}
        </section>

        <section className="step">
          <h2>3. Download</h2>
          {outputs ? (
            <ul className="downloads">
              {outputs.map(output => (
                <li key={output.fileName} className="download-row">
                  <span className="download-label">{output.label}</span>
                  <span className="download-formats">
                    <a className="button" href={output.xlsxUrl} download={`${output.fileName}.xlsx`}>.xlsx</a>
                    <a className="button" href={output.csvUrl} download={`${output.fileName}.csv`}>.csv</a>
                  </span>
                </li>
              ))}
            </ul>
          ) : (
            <p className="hint">Generate pars to download the worksheets.</p>
          )}
        </section>
      </main>
    </>
  )
}

export default App
