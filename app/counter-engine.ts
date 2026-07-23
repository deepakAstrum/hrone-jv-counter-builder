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
  rowOrder: number[];
  reordered: boolean;
};

type PostingSide = "DR" | "CR";

type PreparedEntry = {
  sourceIndex: number;
  amount: number;
  side: PostingSide;
};

type BalancedGroup = {
  debitIds: number[];
  creditIds: number[];
  amount: number;
};

const DEBIT_KEYS = new Set(["40", "29"]);
const CREDIT_KEYS = new Set(["50", "34", "39"]);
const RANDOM_TRIALS = 120_000;
const RANDOM_ATTEMPTS = 20;

export function getPostingSide(value: unknown): PostingSide | "" {
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

function createRandom(seed: number) {
  let state = seed >>> 0 || 0x9e3779b9;
  return () => {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    return (state >>> 0) / 4_294_967_296;
  };
}

function shuffled<T>(values: T[], seed: number) {
  const output = [...values];
  const random = createRandom(seed);
  for (let index = output.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(random() * (index + 1));
    [output[index], output[swapIndex]] = [output[swapIndex], output[index]];
  }
  return output;
}

function walkToTrial(
  candidates: PreparedEntry[],
  seed: number,
  probability: number,
  trial: number,
) {
  const random = createRandom(seed);
  const selected = new Uint8Array(candidates.length);
  let sum = 0;

  for (let index = 0; index < candidates.length; index += 1) {
    if (random() < probability) {
      selected[index] = 1;
      sum += candidates[index].amount;
    }
  }

  for (let index = 0; index < trial; index += 1) {
    const toggleIndex = Math.floor(random() * candidates.length);
    if (selected[toggleIndex]) {
      selected[toggleIndex] = 0;
      sum -= candidates[toggleIndex].amount;
    } else {
      selected[toggleIndex] = 1;
      sum += candidates[toggleIndex].amount;
    }
  }

  return {
    ids: candidates
      .filter((_, index) => selected[index])
      .map((entry) => entry.sourceIndex),
    sum,
  };
}

function findBalancedSubset(
  debitCandidates: PreparedEntry[],
  creditCandidates: PreparedEntry[],
  maxRows: number,
  seed: number,
  debitProbability: number,
  creditProbability: number,
): BalancedGroup | null {
  if (!debitCandidates.length || !creditCandidates.length) return null;

  const debitSeed = seed ^ 0x85ebca6b;
  const creditSeed = seed ^ 0xc2b2ae35;
  const debitRandom = createRandom(debitSeed);
  const debitSelected = new Uint8Array(debitCandidates.length);
  let debitSum = 0;
  let debitCount = 0;

  for (let index = 0; index < debitCandidates.length; index += 1) {
    if (debitRandom() < debitProbability) {
      debitSelected[index] = 1;
      debitSum += debitCandidates[index].amount;
      debitCount += 1;
    }
  }

  const debitStateBySum = new Map<number, number>();
  for (let trial = 0; trial < RANDOM_TRIALS; trial += 1) {
    if (
      debitCount > 0 &&
      debitCount < maxRows &&
      !debitStateBySum.has(debitSum)
    ) {
      debitStateBySum.set(debitSum, trial * 1_000 + debitCount);
    }
    const toggleIndex = Math.floor(debitRandom() * debitCandidates.length);
    if (debitSelected[toggleIndex]) {
      debitSelected[toggleIndex] = 0;
      debitSum -= debitCandidates[toggleIndex].amount;
      debitCount -= 1;
    } else {
      debitSelected[toggleIndex] = 1;
      debitSum += debitCandidates[toggleIndex].amount;
      debitCount += 1;
    }
  }

  const creditRandom = createRandom(creditSeed);
  const creditSelected = new Uint8Array(creditCandidates.length);
  let creditSum = 0;
  let creditCount = 0;

  for (let index = 0; index < creditCandidates.length; index += 1) {
    if (creditRandom() < creditProbability) {
      creditSelected[index] = 1;
      creditSum += creditCandidates[index].amount;
      creditCount += 1;
    }
  }

  for (let trial = 0; trial < RANDOM_TRIALS; trial += 1) {
    const encodedDebitState = debitStateBySum.get(creditSum);
    if (creditCount > 0 && encodedDebitState !== undefined) {
      const debitTrial = Math.floor(encodedDebitState / 1_000);
      const matchedDebitCount = encodedDebitState % 1_000;
      if (creditCount + matchedDebitCount <= maxRows) {
        const debit = walkToTrial(
          debitCandidates,
          debitSeed,
          debitProbability,
          debitTrial,
        );
        const credit = walkToTrial(
          creditCandidates,
          creditSeed,
          creditProbability,
          trial,
        );
        if (debit.sum === credit.sum) {
          return {
            debitIds: debit.ids,
            creditIds: credit.ids,
            amount: debit.sum,
          };
        }
      }
    }

    const toggleIndex = Math.floor(creditRandom() * creditCandidates.length);
    if (creditSelected[toggleIndex]) {
      creditSelected[toggleIndex] = 0;
      creditSum -= creditCandidates[toggleIndex].amount;
      creditCount -= 1;
    } else {
      creditSelected[toggleIndex] = 1;
      creditSum += creditCandidates[toggleIndex].amount;
      creditCount += 1;
    }
  }

  return null;
}

function buildSequentialGroups(
  entries: PreparedEntry[],
  maxRows: number,
): BalancedGroup[] | null {
  const prefix = [0];
  entries.forEach((entry) => {
    const signedAmount = entry.side === "DR" ? entry.amount : -entry.amount;
    prefix.push(prefix[prefix.length - 1] + signedAmount);
  });

  const groups: BalancedGroup[] = [];
  let start = 0;

  while (start < entries.length) {
    const limitExclusive = Math.min(start + maxRows, entries.length);
    let endExclusive = -1;
    for (let candidate = limitExclusive; candidate > start; candidate -= 1) {
      if (prefix[candidate] === prefix[start]) {
        endExclusive = candidate;
        break;
      }
    }
    if (endExclusive === -1) return null;

    const groupEntries = entries.slice(start, endExclusive);
    const amount = groupEntries
      .filter((entry) => entry.side === "DR")
      .reduce((sum, entry) => sum + entry.amount, 0);
    groups.push({
      debitIds: groupEntries
        .filter((entry) => entry.side === "DR")
        .map((entry) => entry.sourceIndex),
      creditIds: groupEntries
        .filter((entry) => entry.side === "CR")
        .map((entry) => entry.sourceIndex),
      amount,
    });
    start = endExclusive;
  }

  return groups;
}

function packBalancedComponents(
  components: BalancedGroup[],
  maxRows: number,
) {
  const groups: BalancedGroup[] = [];
  let current: BalancedGroup = { debitIds: [], creditIds: [], amount: 0 };

  const finishCurrent = () => {
    if (!current.debitIds.length && !current.creditIds.length) return;
    groups.push(current);
    current = { debitIds: [], creditIds: [], amount: 0 };
  };

  components.forEach((component) => {
    const componentRows =
      component.debitIds.length + component.creditIds.length;
    const currentRows = current.debitIds.length + current.creditIds.length;
    if (componentRows > maxRows) {
      throw new Error("A balanced component exceeds the counter row limit.");
    }
    if (currentRows && currentRows + componentRows > maxRows) {
      finishCurrent();
    }
    current.debitIds.push(...component.debitIds);
    current.creditIds.push(...component.creditIds);
    current.amount += component.amount;
  });
  finishCurrent();

  return groups;
}

function buildReorderedGroups(
  entries: PreparedEntry[],
  maxRows: number,
): BalancedGroup[] | null {
  const available = new Set(entries.map((entry) => entry.sourceIndex));
  const components: BalancedGroup[] = [];
  const entryById = new Map(
    entries.map((entry) => [entry.sourceIndex, entry]),
  );

  const availableEntries = (side: PostingSide) =>
    entries.filter(
      (entry) => available.has(entry.sourceIndex) && entry.side === side,
    );

  const reserve = (ids: number[]) => {
    ids.forEach((id) => {
      if (!available.delete(id)) {
        throw new Error("A source row was selected more than once.");
      }
    });
  };

  const debitByAmount = new Map<number, number[]>();
  availableEntries("DR").forEach((entry) => {
    const bucket = debitByAmount.get(entry.amount) ?? [];
    bucket.push(entry.sourceIndex);
    debitByAmount.set(entry.amount, bucket);
  });

  availableEntries("CR").forEach((credit) => {
    const debitId = debitByAmount.get(credit.amount)?.pop();
    if (debitId === undefined) return;
    reserve([debitId, credit.sourceIndex]);
    components.push({
      debitIds: [debitId],
      creditIds: [credit.sourceIndex],
      amount: credit.amount,
    });
  });

  const reserveOneToTwo = (targetSide: PostingSide) => {
    const sourceSide = targetSide === "DR" ? "CR" : "DR";
    const sourceBuckets = new Map<number, number[]>();
    availableEntries(sourceSide).forEach((entry) => {
      const bucket = sourceBuckets.get(entry.amount) ?? [];
      bucket.push(entry.sourceIndex);
      sourceBuckets.set(entry.amount, bucket);
    });
    const sourceAmounts = [...sourceBuckets.keys()].sort(
      (left, right) => left - right,
    );

    availableEntries(targetSide)
      .sort((left, right) => right.amount - left.amount)
      .forEach((target) => {
        if (!available.has(target.sourceIndex)) return;
        let low = 0;
        let high = sourceAmounts.length - 1;

        while (low <= high) {
          const lowAmount = sourceAmounts[low];
          const highAmount = sourceAmounts[high];
          const lowBucket = sourceBuckets.get(lowAmount) ?? [];
          const highBucket = sourceBuckets.get(highAmount) ?? [];

          if (!lowBucket.length) {
            low += 1;
            continue;
          }
          if (!highBucket.length) {
            high -= 1;
            continue;
          }

          const amount = lowAmount + highAmount;
          if (amount < target.amount) {
            low += 1;
            continue;
          }
          if (amount > target.amount) {
            high -= 1;
            continue;
          }
          if (lowAmount === highAmount && lowBucket.length < 2) break;

          const firstSourceId = lowBucket.pop();
          const secondSourceId =
            lowAmount === highAmount ? lowBucket.pop() : highBucket.pop();
          if (
            firstSourceId === undefined ||
            secondSourceId === undefined
          ) {
            break;
          }

          const sourceIds = [firstSourceId, secondSourceId];
          reserve([target.sourceIndex, ...sourceIds]);
          components.push({
            debitIds:
              targetSide === "DR" ? [target.sourceIndex] : sourceIds,
            creditIds:
              targetSide === "CR" ? [target.sourceIndex] : sourceIds,
            amount: target.amount,
          });
          break;
        }
      });
  };

  reserveOneToTwo("DR");
  reserveOneToTwo("CR");

  const groups = packBalancedComponents(components, maxRows);
  let groupIndex = 0;

  while (available.size > maxRows) {
    const debits = availableEntries("DR");
    const credits = availableEntries("CR");
    let match: BalancedGroup | null = null;

    for (let attempt = 0; attempt < RANDOM_ATTEMPTS && !match; attempt += 1) {
      const seed = 0x12345678 + groupIndex * 1_009 + attempt * 9_176;
      const maximumCandidatesPerSide = Math.max(
        1,
        Math.floor(maxRows / 2),
      );
      const debitCandidates = shuffled(debits, seed).slice(
        0,
        Math.min(maximumCandidatesPerSide, debits.length),
      );
      const creditCandidates = shuffled(
        credits,
        seed ^ 0x27d4eb2d,
      ).slice(0, Math.min(maximumCandidatesPerSide, credits.length));
      const debitCandidateTotal = debitCandidates.reduce(
        (sum, entry) => sum + entry.amount,
        0,
      );
      const creditCandidateTotal = creditCandidates.reduce(
        (sum, entry) => sum + entry.amount,
        0,
      );
      if (!debitCandidateTotal && !creditCandidateTotal) continue;

      for (const baseProbability of [0.9, 0.8, 0.7, 0.6, 0.5, 0.4, 0.3]) {
        const targetMean =
          Math.min(debitCandidateTotal, creditCandidateTotal) *
          baseProbability;
        const debitProbability =
          debitCandidateTotal === 0
            ? 0
            : Math.min(0.98, targetMean / debitCandidateTotal);
        const creditProbability =
          creditCandidateTotal === 0
            ? 0
            : Math.min(0.98, targetMean / creditCandidateTotal);

        match = findBalancedSubset(
          debitCandidates,
          creditCandidates,
          maxRows,
          seed,
          debitProbability,
          creditProbability,
        );
        if (match) break;
      }
    }

    if (!match) return null;
    reserve([...match.debitIds, ...match.creditIds]);
    groups.push(match);
    groupIndex += 1;
  }

  if (available.size) {
    const lastEntries = entries.filter((entry) =>
      available.has(entry.sourceIndex),
    );
    const debitIds = lastEntries
      .filter((entry) => entry.side === "DR")
      .map((entry) => entry.sourceIndex);
    const creditIds = lastEntries
      .filter((entry) => entry.side === "CR")
      .map((entry) => entry.sourceIndex);
    const lastDebit = debitIds.reduce(
      (sum, id) => sum + (entryById.get(id)?.amount ?? 0),
      0,
    );
    const lastCredit = creditIds.reduce(
      (sum, id) => sum + (entryById.get(id)?.amount ?? 0),
      0,
    );
    if (lastDebit !== lastCredit) return null;
    groups.push({ debitIds, creditIds, amount: lastDebit });
    reserve([...debitIds, ...creditIds]);
  }

  return available.size === 0 ? groups : null;
}

function completeResult(
  rows: JournalRow[],
  entries: PreparedEntry[],
  groups: BalancedGroup[],
  debitMinor: number,
  creditMinor: number,
  warnings: string[],
): CounterResult {
  const entryById = new Map(
    entries.map((entry) => [entry.sourceIndex, entry]),
  );
  const rowOrder: number[] = [];
  const counters: number[] = [];
  const batches: CounterBatch[] = [];

  groups.forEach((group, index) => {
    const counter = index + 1;
    const groupIds = [
      ...[...group.debitIds].sort((left, right) => left - right),
      ...[...group.creditIds].sort((left, right) => left - right),
    ];
    const startRow = rowOrder.length + 1;
    let batchDebit = 0;
    let batchCredit = 0;

    groupIds.forEach((id) => {
      const entry = entryById.get(id);
      if (!entry) throw new Error("A grouped source row could not be found.");
      rowOrder.push(id);
      counters.push(counter);
      if (entry.side === "DR") batchDebit += entry.amount;
      else batchCredit += entry.amount;
    });

    if (batchDebit !== batchCredit) {
      throw new Error(`Counter ${counter} is not balanced.`);
    }

    batches.push({
      counter,
      startRow,
      endRow: rowOrder.length,
      rowCount: groupIds.length,
      debitMinor: batchDebit,
      creditMinor: batchCredit,
    });
  });

  const uniqueRows = new Set(rowOrder);
  if (rowOrder.length !== rows.length || uniqueRows.size !== rows.length) {
    throw new Error("Source-row integrity check failed.");
  }

  const reordered = rowOrder.some((sourceIndex, index) => sourceIndex !== index);
  const nextWarnings = [...warnings];
  if (reordered) {
    nextWarnings.push(
      "Rows were automatically reordered; every original source row was used exactly once.",
    );
  }

  return {
    status: "complete",
    counters,
    batches,
    debitMinor,
    creditMinor,
    differenceMinor: debitMinor - creditMinor,
    warnings: nextWarnings,
    errors: [],
    blockedAt: null,
    rowOrder,
    reordered,
  };
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
  const emptyCounters: Array<number | null> = Array(rows.length).fill(null);
  const identityOrder = rows.map((_, index) => index);

  if (!postKeyColumn) errors.push("POST_KEY column is required.");
  if (!amountColumn) errors.push("AMOUNT column is required.");
  if (maxRows < 2 || maxRows > 999) {
    errors.push("Maximum rows must be between 2 and 999.");
  }
  if (!rows.length) errors.push("The worksheet has no data rows.");

  if (errors.length || !postKeyColumn || !amountColumn) {
    return {
      status: "invalid",
      counters: emptyCounters,
      batches: [],
      debitMinor: 0,
      creditMinor: 0,
      differenceMinor: 0,
      warnings,
      errors,
      blockedAt: null,
      rowOrder: identityOrder,
      reordered: false,
    };
  }

  let debitMinor = 0;
  let creditMinor = 0;
  let debitDimensionWarnings = 0;
  let creditDimensionWarnings = 0;
  const entries: PreparedEntry[] = [];

  rows.forEach((row, index) => {
    const postKey = String(row[postKeyColumn] ?? "").trim();
    const side = getPostingSide(postKey);
    const amount = toMinorUnits(row[amountColumn]);

    if (!side) {
      errors.push(
        `Data row ${index + 1}: unsupported POST_KEY "${postKey || "(blank)"}".`,
      );
      return;
    }
    if (amount === null || amount < 0) {
      errors.push(
        `Data row ${index + 1}: AMOUNT must be a non-negative number.`,
      );
      return;
    }

    entries.push({ sourceIndex: index, amount, side });
    if (side === "DR") {
      debitMinor += amount;
      const hasCostCenter =
        costCenterColumn && String(row[costCenterColumn] ?? "").trim();
      const hasWbs = wbsColumn && String(row[wbsColumn] ?? "").trim();
      if (!hasCostCenter && !hasWbs) debitDimensionWarnings += 1;
    } else {
      creditMinor += amount;
      const hasProfitCenter =
        profitCenterColumn &&
        String(row[profitCenterColumn] ?? "").trim();
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
      counters: emptyCounters,
      batches: [],
      debitMinor,
      creditMinor,
      differenceMinor: debitMinor - creditMinor,
      warnings,
      errors: errors.slice(0, 20),
      blockedAt: null,
      rowOrder: identityOrder,
      reordered: false,
    };
  }

  if (debitMinor !== creditMinor) {
    return {
      status: "invalid",
      counters: emptyCounters,
      batches: [],
      debitMinor,
      creditMinor,
      differenceMinor: debitMinor - creditMinor,
      warnings,
      errors: [
        "File rejected: total debit and credit amounts do not match. No counters were created.",
      ],
      blockedAt: null,
      rowOrder: identityOrder,
      reordered: false,
    };
  }

  try {
    const sequentialGroups = buildSequentialGroups(entries, maxRows);
    const groups =
      sequentialGroups ?? buildReorderedGroups(entries, maxRows);

    if (!groups) {
      return {
        status: "blocked",
        counters: emptyCounters,
        batches: [],
        debitMinor,
        creditMinor,
        differenceMinor: 0,
        warnings,
        errors: [
          `The balanced file could not be divided into exact counters of ${maxRows} rows or fewer without duplicating or splitting a source row.`,
        ],
        blockedAt: null,
        rowOrder: identityOrder,
        reordered: false,
      };
    }

    return completeResult(
      rows,
      entries,
      groups,
      debitMinor,
      creditMinor,
      warnings,
    );
  } catch (error) {
    return {
      status: "blocked",
      counters: emptyCounters,
      batches: [],
      debitMinor,
      creditMinor,
      differenceMinor: 0,
      warnings,
      errors: [
        error instanceof Error
          ? error.message
          : "Counter construction failed its source-row integrity check.",
      ],
      blockedAt: null,
      rowOrder: identityOrder,
      reordered: false,
    };
  }
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
      WBS_ELE:
        index % 7 === 0
          ? `WBS-DEMO-${String(index % 12).padStart(2, "0")}`
          : "",
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
