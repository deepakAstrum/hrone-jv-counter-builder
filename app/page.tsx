"use client";

import { ChangeEvent, DragEvent, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  buildDemoRows,
  formatAmount,
  getColumn,
  getPostingSide,
  runCounterEngine,
  type CounterResult,
  type JournalRow,
} from "./counter-engine";
import {
  appendJournalOutputColumns,
  getCounterInputIssue,
} from "./workbook-output";

type CellValue = string | number | boolean | Date | null | undefined;

type LoadedFile = {
  fileName: string;
  sheetName: string;
  headers: string[];
  rows: JournalRow[];
  rowNumbers: number[];
  sourceData: ArrayBuffer;
};

type ResultDialog = {
  title: string;
  body: string;
  tone: "error" | "warning";
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
  rowOrder: [],
  reordered: false,
};

function extensionless(fileName: string) {
  return fileName.replace(/\.[^.]+$/, "");
}

function rowsFromSheet(
  sheet: XLSX.WorkSheet,
): Pick<LoadedFile, "headers" | "rows" | "rowNumbers"> {
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

  const rows: JournalRow[] = [];
  const rowNumbers: number[] = [];

  matrix.slice(1).forEach((values, index) => {
    if (!values.some((value) => String(value ?? "").trim() !== "")) return;

    const row: JournalRow = {};
    headers.forEach((header, columnIndex) => {
      row[header] = values[columnIndex] ?? "";
    });
    rows.push(row);
    rowNumbers.push(index + 1);
  });

  return { headers, rows, rowNumbers };
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [loaded, setLoaded] = useState<LoadedFile | null>(null);
  const [result, setResult] = useState<CounterResult>(EMPTY_RESULT);
  const [maxRows, setMaxRows] = useState(999);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [dialog, setDialog] = useState<ResultDialog | null>(null);
  const [message, setMessage] = useState(
    "Upload the monthly HROne extract, or run the 1–980 demonstration.",
  );

  const requiredColumns = useMemo(() => {
    if (!loaded) return [];
    return ["POST_KEY", "AMOUNT"].filter(
      (name) => !getColumn(loaded.headers, name),
    );
  }, [loaded]);

  const counterInputIssue = useMemo(
    () =>
      loaded
        ? getCounterInputIssue(loaded.headers, loaded.rows)
        : null,
    [loaded],
  );
  const hasInputValidationError =
    requiredColumns.length > 0 || Boolean(counterInputIssue);

  const previewRows = useMemo(() => {
    if (!loaded) return [];
    const rowOrder =
      result.status === "complete" &&
      result.rowOrder.length === loaded.rows.length
        ? result.rowOrder
        : loaded.rows.map((_, index) => index);
    return rowOrder.slice(0, 8).map((sourceIndex, outputIndex) => {
      const row = loaded.rows[sourceIndex];
      return {
      line: outputIndex + 1,
      postKey: String(row[getColumn(loaded.headers, "POST_KEY") ?? ""] ?? ""),
      postingSide: getPostingSide(
        row[getColumn(loaded.headers, "POST_KEY") ?? ""],
      ),
      amount: row[getColumn(loaded.headers, "AMOUNT") ?? ""] ?? "",
      costCenter: row[getColumn(loaded.headers, "COSTCENTER") ?? ""] ?? "",
      wbs: row[getColumn(loaded.headers, "WBS_ELE") ?? ""] ?? "",
      profitCenter: row[getColumn(loaded.headers, "PROF_CENT") ?? ""] ?? "",
      counter: result.counters[outputIndex] ?? "",
    };
    });
  }, [loaded, result]);

  function processLoadedFile(next: LoadedFile, nextMaxRows = maxRows) {
    setLoaded(next);
    setResult(EMPTY_RESULT);
    setDialog(null);
    const missing = ["POST_KEY", "AMOUNT"].filter(
      (name) => !getColumn(next.headers, name),
    );
    if (missing.length) {
      setMessage(`Missing required column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}.`);
      setDialog({
        title: "Required column missing",
        body: `Missing required column${missing.length > 1 ? "s" : ""}: ${missing.join(", ")}.`,
        tone: "error",
      });
      return;
    }
    const nextCounterIssue = getCounterInputIssue(next.headers, next.rows);
    if (nextCounterIssue) {
      setMessage(nextCounterIssue.message);
      setDialog({
        title:
          nextCounterIssue.kind === "missing"
            ? "COUNTER column missing"
            : "COUNTER must be blank",
        body: nextCounterIssue.message,
        tone: "error",
      });
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
      const sourceData = await file.arrayBuffer();
      const workbook = XLSX.read(sourceData, {
        type: "array",
        cellDates: false,
        cellStyles: true,
      });
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) throw new Error("No worksheet was found.");
      const parsed = rowsFromSheet(workbook.Sheets[sheetName]);
      processLoadedFile({
        fileName: file.name,
        sheetName,
        sourceData,
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
    const headers = [
      "DOC_DATE",
      "POST_KEY",
      "AMOUNT",
      "COSTCENTER",
      "PROF_CENT",
      "WBS_ELE",
      "COUNTER",
    ];
    const sheet = XLSX.utils.aoa_to_sheet([
      headers,
      ...rows.map((row) => headers.map((header) => row[header] ?? "")),
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Demo");
    const sourceData = XLSX.write(workbook, {
      type: "array",
      bookType: "xlsx",
      cellStyles: true,
    }) as ArrayBuffer;

    processLoadedFile(
      {
        fileName: "balanced-boundary-demo.xlsx",
        sheetName: "Demo",
        headers,
        rows,
        rowNumbers: rows.map((_, index) => index + 1),
        sourceData,
      },
      999,
    );
    setMessage(
      "Demo loaded: row 999 is unbalanced, so the engine steps back to balanced row 980.",
    );
  }

  async function buildCounters() {
    if (!loaded) return;
    if (requiredColumns.length) {
      setDialog({
        title: "Required column missing",
        body: `Missing required column${requiredColumns.length > 1 ? "s" : ""}: ${requiredColumns.join(", ")}.`,
        tone: "error",
      });
      return;
    }
    if (counterInputIssue) {
      setDialog({
        title:
          counterInputIssue.kind === "missing"
            ? "COUNTER column missing"
            : "COUNTER must be blank",
        body: counterInputIssue.message,
        tone: "error",
      });
      return;
    }
    setIsProcessing(true);
    setDialog(null);
    setMessage(
      `Finding exact DR/CR combinations across ${loaded.rows.length.toLocaleString("en-IN")} unique source rows…`,
    );
    await new Promise<void>((resolve) => window.setTimeout(resolve, 20));
    const next = runCounterEngine(loaded.rows, loaded.headers, maxRows);
    setResult(next);
    setIsProcessing(false);
    if (next.status === "complete") {
      setDialog(null);
      setMessage(
        `${next.batches.length} balanced counter batch${next.batches.length === 1 ? "" : "es"} created from ${loaded.rows.length.toLocaleString("en-IN")} unique source rows${next.reordered ? " after automatic reordering" : ""}.`,
      );
    } else if (next.status === "blocked") {
      const body =
        next.errors[0] ??
        `The file could not be divided into balanced counters of ${maxRows} rows or fewer without duplicating or splitting a source row.`;
      setMessage("No exact no-duplication counter combination was found.");
      setDialog({
        title: "Counter creation stopped",
        body,
        tone: "warning",
      });
    } else {
      const balanceMismatch =
        next.errors.some((error) => error.includes("total debit and credit")) &&
        next.differenceMinor !== 0;
      const body = balanceMismatch
        ? `Debit is ${formatAmount(next.debitMinor)} and credit is ${formatAmount(next.creditMinor)}. Difference: ${formatAmount(Math.abs(next.differenceMinor))}. The file was rejected and no counters were created.`
        : next.errors.join(" ");
      setMessage(
        balanceMismatch
          ? "File rejected: debit and credit totals do not match."
          : "Fix the validation errors before creating counters.",
      );
      setDialog({
        title: balanceMismatch ? "Debit and credit do not match" : "File rejected",
        body,
        tone: "error",
      });
    }
  }

  function exportResult() {
    if (
      !loaded ||
      result.status !== "complete" ||
      result.counters.length !== loaded.rows.length
    ) {
      return;
    }

    try {
      const workbook = XLSX.read(loaded.sourceData.slice(0), {
        type: "array",
        cellDates: false,
        cellStyles: true,
      });
      appendJournalOutputColumns({
        workbook,
        sheetName: loaded.sheetName,
        headers: loaded.headers,
        rows: loaded.rows,
        rowNumbers: loaded.rowNumbers,
        counters: result.counters,
        rowOrder: result.rowOrder,
      });
      XLSX.writeFile(
        workbook,
        `${extensionless(loaded.fileName)}-countered.xlsx`,
        { cellStyles: true },
      );
      setMessage(
        `Exported ${loaded.rows.length.toLocaleString("en-IN")} unique original rows with DR/CR appended and the existing COUNTER column populated. Rows were reordered without duplication; dates and row count were preserved.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "The reviewed workbook could not be exported.",
      );
    }
  }

  const balanceOk =
    loaded &&
    result.status !== "idle" &&
    result.debitMinor === result.creditMinor &&
    result.errors.length === 0;
  const canExport =
    Boolean(loaded) &&
    result.status === "complete" &&
    result.counters.length === loaded?.rows.length;

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-mark" aria-hidden="true">
          {/* Shared by the Next.js site and Vite desktop build. */}
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img className="brand-logo" src="./hrone-logo.png" alt="" />
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
          <h2>
            Build counter to make journal data into ERP ready vouchers
            automatically.
          </h2>
          <p className="hero-copy">
            The intelligent app engine first uses the nearest balanced boundary.
            When needed, it automatically combines and reorders original rows
            so every counter remains balanced and contains no more than 999
            lines.
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
          <p className="finding-label">COUNTER CONTROL</p>
          <h3>Every original row is used exactly once.</h3>
          <p>
            Total debit must equal total credit. The app then searches the full
            workbook for exact combinations, reorders those source rows, and
            verifies that nothing was duplicated, dropped, or split.
          </p>
        </div>
        <div className="finding-metrics">
          <span><b>999</b> maximum rows</span>
          <span><b>DR = CR</b> required</span>
          <span><b>0</b> duplicate rows</span>
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
              <p><b>Find exact combinations</b><small>Use original rows from the full workbook only.</small></p>
            </li>
            <li>
              <span>4</span>
              <p><b>Reorder and verify</b><small>Populate COUNTER after the no-duplication check.</small></p>
            </li>
          </ol>

          <button
            className="primary-button"
            type="button"
            disabled={!loaded || hasInputValidationError || isProcessing}
            onClick={() => void buildCounters()}
          >
            {isProcessing ? "Building exact combinations…" : "Build counters"}
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
              {result.status === "complete"
                ? "✓"
                : result.status === "blocked"
                  ? "!"
                  : result.status === "invalid"
                    ? "×"
                    : "·"}
            </span>
            <div>
              <b>
                {result.status === "complete"
                  ? "Ready to export"
                  : result.status === "blocked"
                    ? "Exact combination unavailable"
                    : result.status === "invalid"
                      ? "File rejected"
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
            disabled={!canExport}
            onClick={exportResult}
          >
            Export reviewed workbook
          </button>
        </div>
      </section>

      <section className="data-panel">
        <div className="data-heading">
          <div>
            <p className="kicker">AUDIT PREVIEW</p>
            <h3>First rows and appended output columns</h3>
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
                  <th>DR/CR</th>
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
                    <td><b>{row.postingSide || "—"}</b></td>
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
          <p className="kicker">AUTOMATIC REORDERING</p>
          <h3>Original values move together as complete rows.</h3>
        </div>
        <p>
          The app reorders complete source rows, preserves every date and field,
          keeps the original row count, and rejects the result if the source-row
          identity check is not exact.
        </p>
      </section>

      {dialog && (
        <div className="dialog-backdrop" role="presentation">
          <section
            className={`result-dialog dialog-${dialog.tone}`}
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="result-dialog-title"
          >
            <span className="dialog-mark" aria-hidden="true">
              {dialog.tone === "error" ? "×" : "!"}
            </span>
            <div>
              <p className="kicker">COUNTER BUILDER</p>
              <h3 id="result-dialog-title">{dialog.title}</h3>
              <p>{dialog.body}</p>
              <button
                className="dialog-button"
                type="button"
                autoFocus
                onClick={() => setDialog(null)}
              >
                Close
              </button>
            </div>
          </section>
        </div>
      )}
    </main>
  );
}
