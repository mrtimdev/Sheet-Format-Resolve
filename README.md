# Universal Excel Cleaner

A local web app that removes empty columns or clears matching cells in Excel workbooks. It uses Java 21, Spring Boot 3, Apache POI, Maven, and a Bootstrap 5 interface. The server processes workbooks in memory. The frontend assets are served locally.

## Run locally

Install JDK 21 and Maven 3.9+. From this directory:

```sh
mvn spring-boot:run
```

Open [http://localhost:8080](http://localhost:8080). Upload a workbook, choose a cleaning mode, and select **Clean Excel File**. When the results appear, choose **Download Cleaned Excel**. The page retains the result in memory for download until you reset, select another workbook, or leave the page.

To build and run the standalone JAR:

```sh
mvn verify
java -jar target/excel-cleaner-1.0.0.jar
```

The server binds to loopback by default.

## Cleaning behavior

**Remove empty columns** is the dashboard default. It checks each sheet's data rows within the range of used columns. Missing cells, blanks, whitespace, numeric zero, and text `0` or `NULL` count as empty for this check. A single meaningful value keeps the whole column, including any zeroes stored alongside the data. Formulas, hyperlinks, and comments also keep their columns. Date, ID, Name, Status, Total, and Remarks are protected header labels, matched without case or surrounding whitespace. Set **Header rows** to the number of header rows starting at the first populated row, or use 0 for a sheet without headers.

Surviving columns shift left in their original order. Row positions and the surviving values and styles are preserved. Merged ranges shrink to the surviving cells, with the caption column retained when its merge still spans data or a protected header. Excel recalculates formulas on open. Supported references shift with their columns and a direct reference to a removed column can become `#REF!`.

**Clear matching cells** blanks values in place. Its four options clear empty strings, whitespace-only strings, numeric zero, and the exact text value `"0"`. Already blank and missing cells do not count as changes. The app preserves formulas and their cached values in this mode. When neither mode makes any changes, the response contains the original bytes.

## REST API

Check health:

```sh
curl -fsS http://localhost:8080/api/excel/health
# {"status":"UP"}
```

Remove empty columns with a single header row:

```sh
curl --fail-with-body -sS -D - \
  -F 'file=@report.xlsx' \
  -F 'removeEmptyColumns=true' \
  -F 'headerRows=1' \
  http://localhost:8080/api/excel/clean \
  -o report_cleaned.xlsx
```

Clean a workbook and keep text zero:

```sh
curl --fail-with-body -sS -D - \
  -F 'file=@report.xlsx' \
  -F 'removeEmpty=true' \
  -F 'removeWhitespace=true' \
  -F 'removeNumericZero=true' \
  -F 'removeTextZero=false' \
  http://localhost:8080/api/excel/clean \
  -o report_cleaned.xlsx
```

The dashboard automatically opens a preview when you select a workbook. Choose any worksheet and browse its populated area in pages of 30 rows and 20 columns. The preview shows original row numbers and column letters, including gaps between populated cells. **Auto detect** identifies headers independently for each sheet. It uses worksheet autofilter metadata when present and otherwise looks for recognizable labels or a transition from text labels to numeric, Boolean, or formula data. Text-only data and unusual layouts can be ambiguous. Use **Manual count** to apply one count to every worksheet, with 0 for sheets without headers.

The API defaults to clearing cells for existing callers. Send `removeEmptyColumns=true` to compact columns. In that mode, the four cell-clearing options do not apply. `headerRows` defaults to 1 for existing API clients, accepts 0 through 100 for a manual count, and accepts `-1` for automatic detection. Send `file` and optional zero-based `sheetIndex`, `rowOffset`, and `columnOffset` multipart fields to `POST /api/excel/preview`. Offsets start at the top-left of the selected sheet’s populated area.

The `POST /api/excel/clean` response contains the workbook in its original format. `Content-Disposition` provides a safe download name. Response headers `X-Worksheets-Processed`, `X-Columns-Removed`, `X-Cells-Cleared`, and `X-Processing-Millis` report the run statistics. Workbook responses use `Cache-Control: no-store`.

| HTTP status | Meaning |
| --- | --- |
| 400 | Missing or empty upload, malformed multipart request, or invalid option |
| 413 | Upload or expanded workbook exceeds a resource limit |
| 415 | Unsupported extension or workbook format does not match the extension |
| 422 | Invalid, unsupported, corrupt, encrypted, or sheetless workbook |
| 503 | The workbook processors are busy; retry shortly |
| 500 | The workbook could not be written safely |

Application errors return JSON with a `message` property. The upload cap is 10 MiB, with an 11 MiB multipart request cap and in-memory threshold. OOXML validation caps expansion at 48 MiB collectively, 20 MiB per entry, and 1,024 entries; POI also applies ZIP safety checks. Two workbooks can be processed concurrently. HTTP connection and request-body idle timeouts are 10 and 30 seconds respectively; the browser stops waiting after two minutes.

## Test

```sh
mvn test
```

The integration tests construct representative `.xls` and `.xlsx` workbooks, send them through the REST controller, and reopen the returned files. They cover both cleaning modes, formatting and structure preservation, formula reference updates, multiple header rows, no-op responses, validation errors, and encrypted files.

## Format limits

The app accepts standard OOXML `.xlsx` and Excel 97–2003 BIFF8 `.xls`. It rejects file-open encryption, older BIFF versions, and unsupported formats such as `.xlsb`, `.xlsm`, and CSV. Renaming a file does not convert it. Workbook or worksheet protection that does not encrypt the file is distinct from an opening password.

Column removal updates worksheet autofilter ranges and the positions of filter and sort settings in `.xlsx` workbooks. It returns a 422 if a sheet that needs compaction has an Excel table, pivot table, legacy `.xls` autofilter, drawing, data validation, or array formula that the app cannot safely move. **Clear matching cells** remains available for those workbooks.

When a change is made, Apache POI reserializes the workbook. Advanced Excel features such as digital signatures, VBA in legacy files, ActiveX controls, some pivot or slicer metadata, external connections, and vendor-specific extensions are not guaranteed to round-trip faithfully. Digital signatures may become invalid. Keep the original and check the cleaned copy before using workbooks that depend on those features. Conditional formats are preserved where POI supports them; their displayed results can change when a referenced value changes.
