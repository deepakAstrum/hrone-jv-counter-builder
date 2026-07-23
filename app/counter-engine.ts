export type JournalRow = Record<string, unknown>;

export type CounterBatch = {
  counter: number;
  startRow: number;
  endRow: number;
  rowCount: number;
  debitMinor: number;
  creditMinor: number;
};

export type CounterResult = {
  status: "idle" | "complete" | "blocked" | "invalid";
  counters: Array<number | null>;
  batches: CounterBatch[];
  debitMinor: number;
  creditMinor: number;
  differenceMinor: number;
  warnings: string[];
  errors: string[];
  blockedAt: { startRow: number; limitRow: number } | null;
};

const DEBIT_KEYS = new Set(["40", "29"]);
const CREDIT_KEYS = new Set(["50", "34", "39"]);

export function getPostingSide(value: unknown): "DR" | "CR" | "" {
  const postingKey = String(value ?? "").trim();

  if (DEBIT_KEYS.has(postingKey)) return "DR";
  if (CREDIT_KEYS.has(postingKey)) return "CR";

  return "";
}

export function normaliseColumnName(value: string) {
  return value.toUpperCase().replace(/[^A-Z0-9]/g, "");
}

export function getColumn(headers: string[], expected: string) {
  const target = normaliseColumnName(expected);
  return headers.find((header) => normaliseColumnName(header) === target);
}

function toMinorUnits(value: unknown) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return Math.round(value * 100);
  }
  const cleaned = String(value ?? "")
    .trim()
    .replace(/,/g, "")
    .replace(/[₹$€£\s]/g, "");
  if (!cleaned) return null;
  const parenthesised = cleaned.startsWith("(") && cleaned.endsWith(")");
  const numeric = Number(parenthesised ? cleaned.slice(1, -1) : cleaned);
  if (!Number.isFinite(numeric)) return null;
  return Math.round((parenthesised ? -numeric : numeric) * 100);
}

export function formatAmount(minor: number) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(minor / 100);
}

export function runCounterEngine(
  rows: JournalRow[],
  headers: string[],
  maxRows = 999,
): CounterResult {
  const postKeyColumn = getColumn(headers, "POST_KEY");
  const amountColumn = getColumn(headers, "AMOUNT");
  const costCenterColumn = getColumn(headers, "COSTCENTER");
  const wbsColumn = getColumn(headers, "WBS_ELE");
  const profitCenterColumn = getColumn(headers, "PROF_CENT");
  const errors: string[] = [];
  const warnings: string[] = [];
  const counters: Array<number | null> = Array(rows.length).fill(null);

  if (!postKeyColumn) errors.push("POST_KEY column is required.");
  if (!amountColumn) errors.push("AMOUNT column is required.");
  if (maxRows < 2 || maxRows > 999) errors.push("Maximum rows must be between 2 and 999.");
  if (!rows.length) errors.push("The worksheet has no data rows.");

  if (errors.length || !postKeyColumn || !amountColumn) {
    return {
      status: "invalid",
      counters,
      batches: [],
      debitMinor: 0,
      creditMinor: 0,
      differenceMinor: 0,
      warnings,
      errors,
      blockedAt: null,
    };
  }

  let debitMinor = 0;
  let creditMinor = 0;
  let debitDimensionWarnings = 0;
  let creditDimensionWarnings = 0;
  const signed: number[] = [];

  rows.forEach((row, index) => {
    const postKey = String(row[postKeyColumn] ?? "").trim();
    const amount = toMinorUnits(row[amountColumn]);
    if (!DEBIT_KEYS.has(postKey) && !CREDIT_KEYS.has(postKey)) {
      errors.push(`Data row ${index + 1}: unsupported POST_KEY "${postKey || "(blank)"}".`);
      signed.push(0);
      return;
    }
    if (amount === null || amount < 0) {
      errors.push(`Data row ${index + 1}: AMOUNT must be a non-negative number.`);
      signed.push(0);
      return;
    }

    if (DEBIT_KEYS.has(postKey)) {
      debitMinor += amount;
      signed.push(amount);
      const hasCostCenter = costCenterColumn && String(row[costCenterColumn] ?? "").trim();
      const hasWbs = wbsColumn && String(row[wbsColumn] ?? "").trim();
      if (!hasCostCenter && !hasWbs) debitDimensionWarnings += 1;
    } else {
      creditMinor += amount;
      signed.push(-amount);
      const hasProfitCenter =
        profitCenterColumn && String(row[profitCenterColumn] ?? "").trim();
      if (!hasProfitCenter) creditDimensionWarnings += 1;
    }
  });

  if (debitDimensionWarnings) {
    warnings.push(
      `${debitDimensionWarnings} debit row(s) have neither COSTCENTER nor WBS_ELE.`,
    );
  }
  if (creditDimensionWarnings) {
    warnings.push(
      `${creditDimensionWarnings} credit row(s) have no PROF_CENT.`,
    );
  }

  if (errors.length) {
    return {
      status: "invalid",
      counters,
      batches: [],
      debitMinor,
      creditMinor,
      differenceMinor: debitMinor - creditMinor,
      warnings,
      errors: errors.slice(0, 20),
      blockedAt: null,
    };
  }

  const prefix = [0];
  signed.forEach((value) => prefix.push(prefix[prefix.length - 1] + value));

  const batches: CounterBatch[] = [];
  let start = 0;
  let counter = 1;

  while (start < rows.length) {
    const limitExclusive = Math.min(start + maxRows, rows.length);
    let endExclusive = -1;
    for (let candidate = limitExclusive; candidate > start; candidate -= 1) {
      if (prefix[candidate] === prefix[start]) {
        endExclusive = candidate;
        break;
      }
    }

    if (endExclusive === -1) {
      return {
        status: "blocked",
        counters,
        batches,
        debitMinor,
        creditMinor,
        differenceMinor: debitMinor - creditMinor,
        warnings,
        errors,
        blockedAt: {
          startRow: start + 1,
          limitRow: limitExclusive,
        },
      };
    }

    let batchDebit = 0;
    let batchCredit = 0;
    for (let index = start; index < endExclusive; index += 1) {
      counters[index] = counter;
      if (signed[index] >= 0) batchDebit += signed[index];
      else batchCredit += Math.abs(signed[index]);
    }

    batches.push({
      counter,
      startRow: start + 1,
      endRow: endExclusive,
      rowCount: endExclusive - start,
      debitMinor: batchDebit,
      creditMinor: batchCredit,
    });
    start = endExclusive;
    counter += 1;
  }

  return {
    status: debitMinor === creditMinor ? "complete" : "invalid",
    counters,
    batches,
    debitMinor,
    creditMinor,
    differenceMinor: debitMinor - creditMinor,
    warnings,
    errors:
      debitMinor === creditMinor
        ? []
        : ["The complete file is not balanced; debit and credit totals differ."],
    blockedAt: null,
  };
}

export function buildDemoRows(): JournalRow[] {
  const rows: JournalRow[] = [];
  const base = {
    DOC_DATE: "30-04-2026",
    COUNTER: "",
  };

  for (let index = 0; index < 490; index += 1) {
    rows.push({
      ...base,
      POST_KEY: "40",
      AMOUNT: 100,
      COSTCENTER: `1101${String(index % 40).padStart(5, "0")}`,
      PROF_CENT: "",
      WBS_ELE: index % 7 === 0 ? `WBS-DEMO-${String(index % 12).padStart(2, "0")}` : "",
    });
  }
  for (let index = 0; index < 490; index += 1) {
    rows.push({
      ...base,
      POST_KEY: index % 9 === 0 ? "34" : "50",
      AMOUNT: 100,
      COSTCENTER: "",
      PROF_CENT: `1100${String(index % 8).padStart(3, "0")}`,
      WBS_ELE: "",
    });
  }
  for (let index = 0; index < 10; index += 1) {
    rows.push({
      ...base,
      POST_KEY: "40",
      AMOUNT: 100,
      COSTCENTER: `2200${String(index).padStart(5, "0")}`,
      PROF_CENT: "",
      WBS_ELE: "",
    });
  }
  for (let index = 0; index < 10; index += 1) {
    rows.push({
      ...base,
      POST_KEY: index % 2 ? "39" : "50",
      AMOUNT: 100,
      COSTCENTER: "",
      PROF_CENT: "2200010",
      WBS_ELE: "",
    });
  }
  rows.push({
    ...base,
    POST_KEY: "29",
    AMOUNT: 50,
    COSTCENTER: "220000099",
    PROF_CENT: "",
    WBS_ELE: "",
  });
  rows.push({
    ...base,
    POST_KEY: "50",
    AMOUNT: 50,
    COSTCENTER: "",
    PROF_CENT: "2200010",
    WBS_ELE: "",
  });
  return rows;
}
