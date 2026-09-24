import { useState, type ChangeEvent } from 'react'
import './App.css'
import parWorksheetTemplate from '../templates/par-worksheet.csv?raw'
import prepListTemplate from '../templates/production-prep-list.csv?raw'
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

function parseCsv(content: string): string[][] {
  return Papa.parse<string[]>(content.trim()).data;
}

function hasMovementHeaders(content: string): boolean {
  const headers = Papa.parse<string[]>(content.trim(), { preview: 1 }).data[0] ?? [];
  return Object.entries(MVT_HEADERS).every(
    ([index, name]) => headers[Number(index)]?.trim().toLowerCase() === name.toLowerCase()
  );
}

function App() {
  const [movementReportContent, setMovementReport] = useState('');
  const [errorStatus, setError] = useState(0);
  const [headerWarning, setHeaderWarning] = useState(false);
  const [pwDownload, setPWDownload] = useState<string | null>(null);
  const [plDownload, setPLDownload] = useState<string | null>(null);

  function readMovementReport(event: ChangeEvent<HTMLInputElement>) {
    if (pwDownload) URL.revokeObjectURL(pwDownload);
    setPWDownload(null);
    if (plDownload) URL.revokeObjectURL(plDownload);
    setPLDownload(null);

    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = function () {
        const content = reader.result as string;
        setMovementReport(content);
        setHeaderWarning(!hasMovementHeaders(content));
      };
      reader.readAsText(file);
    }
  }

  function generatePars() {
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

    for (const row of pwRows.slice(1)) {
      const code = row[0]?.trim();
      if (code && weekQtyByCode.has(code)) {
        const weekQty = weekQtyByCode.get(code)!;
        row[4] = String(weekQty);
        const avgPerDay = weekQty / 7;
        row[6] = String(avgPerDay);

        const shelfLife = row[2]?.trim();
        if (shelfLife) row[7] = String(Math.round(avgPerDay * Number(shelfLife)));
      }
    }

    const pwInst = Papa.unparse(pwRows);

    if (pwDownload) URL.revokeObjectURL(pwDownload);
    const pwBlob = new Blob([pwInst], { type: 'text/csv' });
    setPWDownload(URL.createObjectURL(pwBlob));

    const approxParByCode = new Map<string, string>();
    for (const row of pwRows.slice(1)) {
      const code = row[0]?.trim();
      if (code) approxParByCode.set(code, row[7]);
    }

    const plRows = parseCsv(prepListTemplate);
    for (const row of plRows.slice(3)) {
      for (const [pluIndex, parIndex] of [[1, 3], [9, 11]] as const) {
        const code = row[pluIndex]?.trim();
        if (code && approxParByCode.has(code)) {
          row[parIndex] = approxParByCode.get(code)!;
        } else if (code && weekQtyByCode.has(code)) {
          // not on the par worksheet: using a one-week par in absence of shelf life value
          row[parIndex] = String(weekQtyByCode.get(code));
        }
      }
    }

    const plInst = Papa.unparse(plRows);

    if (plDownload) URL.revokeObjectURL(plDownload);
    const plBlob = new Blob([plInst], { type: 'text/csv' });
    setPLDownload(URL.createObjectURL(plBlob));
  }

  return (
    <main className="card">
      <h1>Production Par</h1>

      <section className="step">
        <h2>1. Upload Movement Report</h2>
        <label htmlFor="fileInput">
          <input type="file" id="fileInput" accept=".csv" onChange={readMovementReport} />
        </label>
        {headerWarning && (
          <p className="warning">This document doesn't appear to have the correct headers. Outputs may not generate correctly</p>
        )}
      </section>

      <section className="step">
        <h2>2. Generate</h2>
        <button className="button" onClick={generatePars}>Generate Pars</button>
        {errorStatus !== 0 && <p className="error">No file inputted</p>}
      </section>

      <section className="step">
        <h2>3. Download</h2>
        {pwDownload ? (
          <div className="downloads">
            <a className="button" href={pwDownload} download="par-worksheet.csv">Download Par Worksheet</a>
            {plDownload && (
              <a className="button" href={plDownload} download="production-prep-list.csv">Download Production Prep List</a>
            )}
          </div>
        ) : (
          <p className="hint">Generate pars to download the worksheets.</p>
        )}
      </section>
    </main>
  )
}

export default App
