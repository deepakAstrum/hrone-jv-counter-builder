# HROne Counter Builder

Browser-based validator and counter generator for monthly HROne-to-SAP journal
extracts.

The implemented rule assigns posting keys `40` and `29` to debit, and `50`,
`34`, and `39` to credit. For each counter it examines at most 999 sequential
data rows and chooses the latest preceding row where cumulative debit equals
cumulative credit. If no balanced boundary exists, processing stops and leaves
the remaining counters blank.

The app accepts `.xls`, `.xlsx`, and `.csv`, processes the file locally in the
browser, previews the result, and exports an `.xlsx` with `COUNTER` as the last
column.
