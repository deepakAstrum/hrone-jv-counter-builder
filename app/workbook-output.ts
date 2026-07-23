import * as XLSX from "xlsx";

import {
  getColumn,
  getPostingSide,
  type JournalRow,
} from "./counter-engine";

export const POSTING_SIDE_HEADER = "DR/CR";
export const COUNTER_BUILDER_HEADER = "COUNTER BUILDER";

type AppendOutputColumnsInput = {
  workbook: XLSX.WorkBook;
  sheetName: string;
  headers: string[];
  rows: JournalRow[];
  rowNumbers: number[];
  counters: Array<number | null>;
};

export function appendJournalOutputColumns({
  workbook,
  sheetName,
  headers,
  rows,
  rowNumbers,
  counters,
}: AppendOutputColumnsInput) {
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error(`Worksheet "${sheetName}" could not be found.`);
  }

  if (!sheet["!ref"]) {
    throw new Error("The selected worksheet does not contain any cells.");
  }

  if (rows.length !== rowNumbers.length || rows.length !== counters.length) {
    throw new Error("The output rows no longer match the source worksheet.");
  }

  const postingKeyColumn =
    getColumn(headers, "POST_KEY") ??
    getColumn(headers, "POSTING_KEY") ??
    getColumn(headers, "POSTINGKEY");

  if (!postingKeyColumn) {
    throw new Error("POST_KEY column was not found.");
  }

  const sourceRange = XLSX.utils.decode_range(sheet["!ref"]);
  const postingSideColumn = headers.length;
  const counterBuilderColumn = headers.length + 1;

  const setTextCell = (row: number, column: number, value: string) => {
    sheet[XLSX.utils.encode_cell({ r: row, c: column })] = {
      t: "s",
      v: value,
    };
  };

  const setOutputCell = (
    row: number,
    column: number,
    value: string | number,
  ) => {
    sheet[XLSX.utils.encode_cell({ r: row, c: column })] =
      typeof value === "number"
        ? { t: "n", v: value }
        : { t: "s", v: value };
  };

  setTextCell(0, postingSideColumn, POSTING_SIDE_HEADER);
  setTextCell(0, counterBuilderColumn, COUNTER_BUILDER_HEADER);

  rows.forEach((row, index) => {
    const worksheetRow = rowNumbers[index];

    if (worksheetRow < 1 || worksheetRow > sourceRange.e.r) {
      throw new Error("A source row falls outside the original worksheet.");
    }

    setOutputCell(
      worksheetRow,
      postingSideColumn,
      getPostingSide(row[postingKeyColumn]),
    );
    setOutputCell(
      worksheetRow,
      counterBuilderColumn,
      counters[index] ?? "",
    );
  });

  sourceRange.e.c = Math.max(sourceRange.e.c, counterBuilderColumn);
  sheet["!ref"] = XLSX.utils.encode_range(sourceRange);

  return {
    postingSideColumn,
    counterBuilderColumn,
    sourceLastRow: sourceRange.e.r,
  };
}
