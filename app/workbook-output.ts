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
  rowOrder: number[];
};

export function appendJournalOutputColumns({
  workbook,
  sheetName,
  headers,
  rows,
  rowNumbers,
  counters,
  rowOrder,
}: AppendOutputColumnsInput) {
  const sheet = workbook.Sheets[sheetName];

  if (!sheet) {
    throw new Error(`Worksheet "${sheetName}" could not be found.`);
  }

  if (!sheet["!ref"]) {
    throw new Error("The selected worksheet does not contain any cells.");
  }

  if (
    rows.length !== rowNumbers.length ||
    rows.length !== counters.length ||
    rows.length !== rowOrder.length
  ) {
    throw new Error("The output rows no longer match the source worksheet.");
  }
  if (
    new Set(rowOrder).size !== rows.length ||
    rowOrder.some(
      (sourceIndex) => sourceIndex < 0 || sourceIndex >= rows.length,
    )
  ) {
    throw new Error(
      "Source-row integrity check failed: rows would be duplicated or omitted.",
    );
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
  const sourceCells = rowNumbers.map((worksheetRow) =>
    headers.map((_, column) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: worksheetRow, c: column })];
      return cell ? structuredClone(cell) : undefined;
    }),
  );
  const sourceRowStyles = rowNumbers.map((worksheetRow) => {
    const rowStyle = sheet["!rows"]?.[worksheetRow];
    return rowStyle ? structuredClone(rowStyle) : undefined;
  });

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
  const columns = sheet["!cols"] ?? [];
  columns[postingSideColumn] = {
    ...(columns[postingSideColumn] ?? {}),
    wch: Math.max(columns[postingSideColumn]?.wch ?? 0, 10),
  };
  columns[counterBuilderColumn] = {
    ...(columns[counterBuilderColumn] ?? {}),
    wch: Math.max(columns[counterBuilderColumn]?.wch ?? 0, 18),
  };
  sheet["!cols"] = columns;

  rowOrder.forEach((sourceIndex, outputIndex) => {
    const worksheetRow = rowNumbers[outputIndex];

    if (worksheetRow < 1 || worksheetRow > sourceRange.e.r) {
      throw new Error("A source row falls outside the original worksheet.");
    }

    headers.forEach((_, column) => {
      const address = XLSX.utils.encode_cell({ r: worksheetRow, c: column });
      const sourceCell = sourceCells[sourceIndex][column];
      if (sourceCell) sheet[address] = sourceCell;
      else delete sheet[address];
    });

    if (sheet["!rows"]) {
      const sourceRowStyle = sourceRowStyles[sourceIndex];
      if (sourceRowStyle) sheet["!rows"][worksheetRow] = sourceRowStyle;
      else delete sheet["!rows"][worksheetRow];
    }

    const row = rows[sourceIndex];
    setOutputCell(
      worksheetRow,
      postingSideColumn,
      getPostingSide(row[postingKeyColumn]),
    );
    setOutputCell(
      worksheetRow,
      counterBuilderColumn,
      counters[outputIndex] ?? "",
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
