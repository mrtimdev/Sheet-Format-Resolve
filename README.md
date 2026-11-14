# Universal Excel Cleaner

A local web app that clears unwanted values from every worksheet in an Excel workbook. It uses Java 21, Spring Boot 3, Apache POI, Maven, and a Bootstrap 5 interface. The server processes workbooks in memory; it uses no database, accounts, external services, or upload/output directories. The frontend assets are served locally.

## Run locally

Install JDK 21 and Maven 3.9+. From this directory:

```sh
mvn spring-boot:run
```

Open [http://localhost:8080](http://localhost:8080). Drop or browse for a workbook, select the cleaning rules, and choose **Clean Excel File**. When the results appear, choose **Download Cleaned Excel**. The browser saves the download normally. The page retains the result in memory for download until you reset, select another workbook, or leave the page.

To build and run the standalone JAR:

```sh
mvn verify
java -jar target/excel-cleaner-1.0.0.jar
```

The server binds to loopback by default.

## Cleaning behavior

All four options default to enabled. They independently clear stored empty strings, nonempty strings made only of whitespace, numeric zero, and the exact text value `"0"`. Numeric zero refers to the underlying Excel value, regardless of number formatting. Already blank cells and missing cells are left alone and do not count as changes. Values like `" 0 "`, booleans, errors, and formulas are never cleared. Formula caches are left as read; Excel may update their results on recalculation.

Cells are blanked in place. The app does not remove rows, columns, or sheets. Tests verify cell styles, hyperlinks, merged ranges, hidden rows/columns/sheets, freeze panes, row heights, column widths, conditional formatting rules, and zero-valued formulas. If no cells change, the response contains the original workbook bytes.

## REST API

Check health:

```sh
curl -fsS http://localhost:8080/api/excel/health
# {"status":"UP"}
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

The `POST /api/excel/clean` response contains the workbook in its original format. `Content-Disposition` provides a safe download name. Response headers `X-Worksheets-Processed`, `X-Cells-Cleared`, and `X-Processing-Millis` report the run statistics. Workbook responses use `Cache-Control: no-store`.

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

The integration tests construct representative `.xls` and `.xlsx` workbooks, send them through the REST controller, and reopen the returned files. They cover the cleaning rules, formatting and structure preservation, option toggles, no-op responses, validation errors, and encrypted files.

## Format limits

The app accepts standard OOXML `.xlsx` and Excel 97–2003 BIFF8 `.xls`. It rejects file-open encryption, older BIFF versions, and unsupported formats such as `.xlsb`, `.xlsm`, and CSV. Renaming a file does not convert it. Workbook or worksheet protection that does not encrypt the file is distinct from an opening password.

When a change is made, Apache POI reserializes the workbook. Advanced Excel features such as digital signatures, VBA in legacy files, ActiveX controls, some pivot or slicer metadata, external connections, and vendor-specific extensions are not guaranteed to round-trip faithfully. Digital signatures may become invalid. Keep the original and check the cleaned copy before using workbooks that depend on those features. Conditional formats are preserved where POI supports them; their displayed results can change when a referenced value changes.
