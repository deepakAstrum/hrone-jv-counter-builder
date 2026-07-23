"use client";

import { ChangeEvent, DragEvent, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  buildDemoRows,
  formatAmount,
  getColumn,
  normaliseColumnName,
  runCounterEngine,
  type CounterResult,
  type JournalRow,
} from "./counter-engine";

type CellValue = string | number | boolean | Date | null | undefined;

type LoadedFile = {
  fileName: string;
  sheetName: string;
  headers: string[];
  rows: JournalRow[];
};

const EMPTY_RESULT: CounterResult = {
  status: "idle",
  counters: [],
  batches: [],
  debitMinor: 0,
  creditMinor: 0,
  differenceMinor: 0,
  warnings: [],
  errors: [],
  blockedAt: null,
};

function extensionless(fileName: string) {
  return fileName.replace(/\.[^.]+$/, "");
}

function rowsFromSheet(sheet: XLSX.WorkSheet): Pick<LoadedFile, "headers" | "rows"> {
  const matrix = XLSX.utils.sheet_to_json<CellValue[]>(sheet, {
    header: 1,
    defval: "",
    raw: true,
  });

  if (!matrix.length) {
    throw new Error("The selected worksheet is empty.");
  }

  const headers = matrix[0].map((value, index) => {
    const label = String(value ?? "").trim();
    return label || `COLUMN_${index + 1}`;
  });

  const rows = matrix
    .slice(1)
    .filter((row) => row.some((value) => String(value ?? "").trim() !== ""))
    .map((values) => {
      const row: JournalRow = {};
      headers.forEach((header, index) => {
        row[header] = values[index] ?? "";
      });
      return row;
    });

  return { headers, rows };
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [loaded, setLoaded] = useState<LoadedFile | null>(null);
  const [result, setResult] = useState<CounterResult>(EMPTY_RESULT);
  const [maxRows, setMaxRows] = useState(999);
  const [isDragging, setIsDragging] = useState(false);
  const [message, setMessage] = useState(
    "Upload the monthly HROne extract, or run the 1–980 demonstration.",
  );

  const requiredColumns = useMemo(() => {
    if (!loaded) return [];
    return ["POST_KEY", "AMOUNT"].filter(
      (name) => !getColumn(loaded.headers, name),
    );
  }, [loaded]);

  const previewRows = useMemo(() => {
    if (!loaded) return [];
    return loaded.rows.slice(0, 8).map((row, index) => ({
      line: index + 1,
      postKey: String(row[getColumn(loaded.headers, "POST_KEY") ?? ""] ?? ""),
      amount: row[getColumn(loaded.headers, "AMOUNT") ?? ""] ?? "",
      costCenter: row[getColumn(loaded.headers, "COSTCENTER") ?? ""] ?? "",
      wbs: row[getColumn(loaded.headers, "WBS_ELE") ?? ""] ?? "",
      profitCenter: row[getColumn(loaded.headers, "PROF_CENT") ?? ""] ?? "",
      counter: result.counters[index] ?? "",
    }));
  }, [loaded, result]);

  function processLoadedFile(next: LoadedFile, nextMaxRows = maxRows) {
    setLoaded(next);
    setResult(EMPTY_RESULT);
    const missing = ["POST_KEY", "AMOUNT"].filter(
      (name) => !getColumn(next.headers, name),
    );
    if (missing.length) {
      setMessage(`Missing required column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}.`);
      return;
    }
    setMessage(
      `${next.rows.length.toLocaleString("en-IN")} rows loaded from ${next.sheetName}. Ready to build counters.`,
    );
    if (nextMaxRows !== maxRows) setMaxRows(nextMaxRows);
  }

  async function readFile(file: File) {
    try {
      setMessage(`Reading ${file.name}…`);
      const workbook = XLSX.read(await file.arrayBuffer(), {
        type: "array",
        cellDates: true,
      });
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) throw new Error("No worksheet was found.");
      const parsed = rowsFromSheet(workbook.Sheets[sheetName]);
      processLoadedFile({
        fileName: file.name,
        sheetName,
        ...parsed,
      });
    } catch (error) {
      setLoaded(null);
      setResult(EMPTY_RESULT);
      setMessage(
        error instanceof Error
          ? error.message
          : "The workbook could not be read. Please use .xls, .xlsx, or .csv.",
      );
    }
  }

  function onFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (file) void readFile(file);
  }

  function onDrop(event: DragEvent<HTMLButtonElement>) {
    event.preventDefault();
    setIsDragging(false);
    const file = event.dataTransfer.files?.[0];
    if (file) void readFile(file);
  }

  function loadDemo() {
    const rows = buildDemoRows();
    processLoadedFile(
      {
        fileName: "balanced-boundary-demo.xlsx",
        sheetName: "Demo",
        headers: [
          "DOC_DATE",
          "POST_KEY",
          "AMOUNT",
          "COSTCENTER",
          "PROF_CENT",
          "WBS_ELE",
          "COUNTER",
        ],
        rows,
      },
      999,
    );
    setMessage(
      "Demo loaded: row 999 is unbalanced, so the engine steps back to balanced row 980.",
    );
  }

  function buildCounters() {
    if (!loaded || requiredColumns.length) return;
    const next = runCounterEngine(loaded.rows, loaded.headers, maxRows);
    setResult(next);
    if (next.status === "complete") {
      setMessage(
        `${next.batches.length} balanced counter batch${next.batches.length === 1 ? "" : "es"} created.`,
      );
    } else if (next.status === "blocked") {
      setMessage(
        "Counter creation stopped: no balanced boundary exists inside the allowed row window.",
      );
    } else {
      setMessage("Fix the validation errors before creating counters.");
    }
  }

  function exportResult() {
    if (!loaded || result.status !== "complete") return;
    const counterHeader = getColumn(loaded.headers, "COUNTER");
    const exportHeaders = loaded.headers.filter(
      (header) => normaliseColumnName(header) !== "COUNTER",
    );
    exportHeaders.push(counterHeader || "COUNTER");

    const matrix: CellValue[][] = [
      exportHeaders,
      ...loaded.rows.map((row, index) =>
        exportHeaders.map((header) =>
          normaliseColumnName(header) === "COUNTER"
            ? result.counters[index]
            : row[header] ?? "",
        ),
      ),
    ];

    const sheet = XLSX.utils.aoa_to_sheet(matrix);
    sheet["!autofilter"] = { ref: `A1:${XLSX.utils.encode_col(exportHeaders.length - 1)}1` };
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, loaded.sheetName.slice(0, 31));
    XLSX.writeFile(workbook, `${extensionless(loaded.fileName)}-countered.xlsx`);
  }

  const balanceOk =
    loaded &&
    result.debitMinor === result.creditMinor &&
    result.errors.length === 0;

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-mark" aria-hidden="true">
          HR
        </div>
        <div className="brand-copy">
          <p>HROne → SAP</p>
          <h1>Counter Builder</h1>
        </div>
        <div className="topbar-status">
          <span className="status-dot" />
          Monthly journal control
        </div>
      </header>

      <section className="hero">
        <div>
          <p className="eyebrow">BALANCED DOCUMENT SPLITTING</p>
          <h2>Build every counter on a true Dr = Cr boundary.</h2>
          <p className="hero-copy">
            The engine checks up to 999 lines, steps backward to the nearest
            balanced line, assigns one counter, then repeats from the next row.
          </p>
        </div>
        <div className="hero-badge">
          <span>Posting keys</span>
          <strong>40 · 29</strong>
          <small>Debit</small>
          <strong>50 · 34 · 39</strong>
          <small>Credit</small>
        </div>
      </section>

      <section className="finding-card" aria-label="Workbook finding">
        <div className="finding-icon">!</div>
        <div>
          <p className="finding-label">JV_Jain.xls finding</p>
          <h3>The file balances overall, but not within the first 999 rows.</h3>
          <p>
            7,942 rows · Dr ₹366,823,342.29 · Cr ₹366,823,342.29 · first
            balanced prefix at row 7,942. Strict sequential countering therefore
            stops at counter 1 until a grouping or reordering rule is supplied.
          </p>
        </div>
        <div className="finding-metrics">
          <span><b>43</b> profit centres</span>
          <span><b>206</b> cost centres</span>
          <span><b>49</b> WBS elements</span>
        </div>
      </section>

      <section className="workspace-grid">
        <div className="panel upload-panel">
          <div className="panel-heading">
            <span className="step-number">01</span>
            <div>
              <p className="kicker">SOURCE</p>
              <h3>Load monthly extract</h3>
            </div>
          </div>

          <input
            ref={inputRef}
            className="sr-only"
            type="file"
            accept=".xls,.xlsx,.csv"
            onChange={onFileChange}
          />
          <button
            className={`drop-zone ${isDragging ? "is-dragging" : ""}`}
            type="button"
            onClick={() => inputRef.current?.click()}
            onDragEnter={() => setIsDragging(true)}
            onDragLeave={() => setIsDragging(false)}
            onDragOver={(event) => event.preventDefault()}
            onDrop={onDrop}
          >
            <span className="upload-glyph">↑</span>
            <strong>{loaded ? loaded.fileName : "Drop Excel extract here"}</strong>
            <small>
              {loaded
                ? `${loaded.sheetName} · ${loaded.rows.length.toLocaleString("en-IN")} rows`
                : ".xls, .xlsx, or .csv · processed in your browser"}
            </small>
          </button>
          <button className="text-button" type="button" onClick={loadDemo}>
            Run the 1–980 boundary demonstration
          </button>

          <div className="field-map">
            <div>
              <span className="map-side debit">DR</span>
              <p><b>40, 29</b><small>COSTCENTER or WBS_ELE</small></p>
            </div>
            <div>
              <span className="map-side credit">CR</span>
              <p><b>50, 34, 39</b><small>PROF_CENT</small></p>
            </div>
          </div>
        </div>

        <div className="panel rules-panel">
          <div className="panel-heading">
            <span className="step-number">02</span>
            <div>
              <p className="kicker">CONTROL</p>
              <h3>Apply counter rule</h3>
            </div>
          </div>

          <label className="max-lines">
            <span>Maximum lines per counter</span>
            <input
              type="number"
              min={2}
              max={999}
              value={maxRows}
              onChange={(event) =>
                setMaxRows(
                  Math.max(2, Math.min(999, Number(event.target.value) || 999)),
                )
              }
            />
          </label>

          <ol className="rule-list">
            <li>
              <span>1</span>
              <p><b>Read forward</b><small>Open a window of no more than {maxRows} lines.</small></p>
            </li>
            <li>
              <span>2</span>
              <p><b>Test the balance</b><small>Dr keys add; Cr keys subtract.</small></p>
            </li>
            <li>
              <span>3</span>
              <p><b>Step backward</b><small>Choose the latest preceding line where net = 0.</small></p>
            </li>
            <li>
              <span>4</span>
              <p><b>Assign and repeat</b><small>Write COUNTER, then start at the next unassigned row.</small></p>
            </li>
          </ol>

          <button
            className="primary-button"
            type="button"
            disabled={!loaded || requiredColumns.length > 0}
            onClick={buildCounters}
          >
            Build counters
            <span>→</span>
          </button>
        </div>

        <div className="panel result-panel">
          <div className="panel-heading">
            <span className="step-number">03</span>
            <div>
              <p className="kicker">RESULT</p>
              <h3>Verify and export</h3>
            </div>
          </div>

          <div className={`result-state state-${result.status}`}>
            <span className="state-symbol">
              {result.status === "complete" ? "✓" : result.status === "blocked" ? "!" : "·"}
            </span>
            <div>
              <b>
                {result.status === "complete"
                  ? "Ready to export"
                  : result.status === "blocked"
                    ? "Linking rule required"
                    : "Waiting for run"}
              </b>
              <small>{message}</small>
            </div>
          </div>

          <div className="totals-grid">
            <div><span>Debit</span><b>{formatAmount(result.debitMinor)}</b></div>
            <div><span>Credit</span><b>{formatAmount(result.creditMinor)}</b></div>
            <div className={balanceOk ? "balanced" : ""}>
              <span>Difference</span>
              <b>{formatAmount(Math.abs(result.differenceMinor))}</b>
            </div>
            <div><span>Counters</span><b>{result.batches.length}</b></div>
          </div>

          {result.blockedAt && (
            <div className="blocker-note">
              No zero-balance boundary from data row{" "}
              <b>{result.blockedAt.startRow.toLocaleString("en-IN")}</b> to{" "}
              <b>{result.blockedAt.limitRow.toLocaleString("en-IN")}</b>.
              Unassigned rows remain blank.
            </div>
          )}

          <button
            className="export-button"
            type="button"
            disabled={result.status !== "complete"}
            onClick={exportResult}
          >
            Export countered workbook
          </button>
        </div>
      </section>

      <section className="data-panel">
        <div className="data-heading">
          <div>
            <p className="kicker">AUDIT PREVIEW</p>
            <h3>First rows and assigned counter</h3>
          </div>
          <p>{message}</p>
        </div>

        {previewRows.length ? (
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Line</th>
                  <th>Post key</th>
                  <th>Amount</th>
                  <th>Cost center</th>
                  <th>WBS element</th>
                  <th>Profit center</th>
                  <th>Counter</th>
                </tr>
              </thead>
              <tbody>
                {previewRows.map((row) => (
                  <tr key={row.line}>
                    <td>{row.line}</td>
                    <td><span className={`key-pill ${["40", "29"].includes(row.postKey) ? "dr" : "cr"}`}>{row.postKey}</span></td>
                    <td>{String(row.amount)}</td>
                    <td>{String(row.costCenter) || "—"}</td>
                    <td>{String(row.wbs) || "—"}</td>
                    <td>{String(row.profitCenter) || "—"}</td>
                    <td><b>{row.counter || "—"}</b></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="empty-table">
            <span>↳</span>
            <p>Journal rows will appear here after you load an extract.</p>
          </div>
        )}
      </section>

      <section className="next-rule">
        <div>
          <p className="kicker">REQUIRED FOR JV_JAIN.XLS</p>
          <h3>Add one auditable grouping key before automatic reordering.</h3>
        </div>
        <p>
          Debit rows carry COSTCENTER or WBS_ELE, while credit rows carry
          PROF_CENT. Because the extract has no common linking field, the app
          must not guess which debits belong to which profit centre. The next
          stage should load a maintained Cost Center/WBS → Profit Center mapping,
          form balanced profit-centre blocks, and then pack those blocks into
          counters of no more than 999 lines.
        </p>
      </section>
    </main>
  );
}
