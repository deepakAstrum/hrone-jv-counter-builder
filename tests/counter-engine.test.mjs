import assert from "node:assert/strict";
import test from "node:test";

import {
  getMinimumBalancedGroupRows,
  runCounterEngine,
} from "../app/counter-engine.ts";

const headers = ["POST_KEY", "AMOUNT", "COUNTER"];

function journalRows(postingKey, amount, count) {
  return Array.from({ length: count }, () => ({
    POST_KEY: postingKey,
    AMOUNT: amount,
    COUNTER: "",
  }));
}

test("requires a final-fit group when the remainder is close to the row limit", () => {
  assert.equal(getMinimumBalancedGroupRows(4_503, 999), 2);
  assert.equal(getMinimumBalancedGroupRows(1_995, 999), 2);
  assert.equal(getMinimumBalancedGroupRows(1_501, 999), 502);
  assert.equal(getMinimumBalancedGroupRows(1_006, 999), 7);
  assert.equal(getMinimumBalancedGroupRows(999, 999), 0);
});

test("builds exact counters without duplicating or omitting source rows", () => {
  const rows = [
    ...journalRows("40", 3, 400),
    ...journalRows("50", 2, 600),
  ];

  const result = runCounterEngine(rows, headers, 999);

  assert.equal(result.status, "complete");
  assert.equal(result.debitMinor, result.creditMinor);
  assert.equal(result.rowOrder.length, rows.length);
  assert.equal(new Set(result.rowOrder).size, rows.length);
  assert.ok(result.batches.every((batch) => batch.rowCount <= 999));
  assert.ok(
    result.batches.every(
      (batch) => batch.debitMinor === batch.creditMinor,
    ),
  );
});

test("still rejects a workbook whose debit and credit totals differ", () => {
  const rows = [
    ...journalRows("40", 100, 2),
    ...journalRows("50", 100, 1),
  ];

  const result = runCounterEngine(rows, headers, 999);

  assert.equal(result.status, "invalid");
  assert.match(result.errors[0], /total debit and credit amounts do not match/i);
});
