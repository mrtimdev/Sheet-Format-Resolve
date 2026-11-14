package com.example.excelcleaner;

import static org.assertj.core.api.Assertions.assertThat;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.http.ContentDisposition;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;

@SpringBootTest
@AutoConfigureMockMvc
class ExcelCleanerIntegrationTest {
    private static final String API = "/api/excel/clean";
    @Autowired MockMvc mvc;

    @Test
    void previewsDataBeyondTheFirstPageOfRowsAndColumns() throws Exception {
        String detail = "รายละเอียด ".repeat(40);
        byte[] input;
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Wide report");
            sheet.createRow(1).createCell(3).setCellValue("Start");
            sheet.createRow(34).createCell(24).setCellValue(detail);
            input = WorkbookTestData.bytes(workbook);
        }
        var upload = file("wide.xlsx", input);
        mvc.perform(multipart("/api/excel/preview").file(upload))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.totalRows").value(34))
                .andExpect(jsonPath("$.totalColumns").value(22))
                .andExpect(jsonPath("$.moreRows").value(true))
                .andExpect(jsonPath("$.moreColumns").value(true));
        mvc.perform(multipart("/api/excel/preview").file(upload)
                        .param("rowOffset", "30").param("columnOffset", "20"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.columns[1]").value("Y"))
                .andExpect(jsonPath("$.rows[0].rowNumber").value(32))
                .andExpect(jsonPath("$.rows[3].rowNumber").value(35))
                .andExpect(jsonPath("$.rows[3].values[1]").value(detail))
                .andExpect(jsonPath("$.moreRows").value(false))
                .andExpect(jsonPath("$.moreColumns").value(false));
        mvc.perform(multipart("/api/excel/preview").file(upload).param("rowOffset", "34"))
                .andExpect(status().isBadRequest());
    }

    @ParameterizedTest
    @ValueSource(strings = {"xlsx", "xls"})
    void previewAndCleaningAgreeOnHeadersForEachWorksheet(String extension) throws Exception {
        byte[] input;
        try (Workbook workbook = extension.equals("xls") ? new HSSFWorkbook() : new XSSFWorkbook()) {
            var report = workbook.createSheet("Report");
            report.createRow(0).createCell(0).setCellValue(" \u00a0");
            report.createRow(2).createCell(1).setCellValue("Dispatch Report");
            report.addMergedRegion(CellRangeAddress.valueOf("B3:D3"));
            var labels = report.createRow(3);
            labels.createCell(1).setCellValue("Date");
            labels.createCell(2).setCellValue("Unused");
            labels.createCell(3).setCellValue("Quantity");
            report.createRow(4).createCell(3).setCellValue(20);
            var raw = workbook.createSheet("Raw");
            raw.createRow(0).createCell(0).setCellValue(100);
            raw.getRow(0).createCell(1).setCellValue(0);
            workbook.createSheet("Blank");
            input = WorkbookTestData.bytes(workbook);
        }
        var upload = file("preview." + extension, input);
        mvc.perform(multipart("/api/excel/preview").file(upload))
                .andExpect(status().isOk())
                .andExpect(header().string("Cache-Control", "no-store"))
                .andExpect(jsonPath("$.sheetNames[1]").value("Raw"))
                .andExpect(jsonPath("$.firstPopulatedRow").value(3))
                .andExpect(jsonPath("$.detectedHeaderRows").value(2))
                .andExpect(jsonPath("$.columns[0]").value("B"))
                .andExpect(jsonPath("$.rows[1].values[1]").value("Unused"));
        mvc.perform(multipart("/api/excel/preview").file(upload).param("sheetIndex", "1"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.detectedHeaderRows").value(0))
                .andExpect(jsonPath("$.rows[0].values[0]").value("100"));
        mvc.perform(multipart("/api/excel/preview").file(upload).param("sheetIndex", "2"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.rows").isEmpty());
        mvc.perform(multipart("/api/excel/preview").file(upload).param("sheetIndex", "3"))
                .andExpect(status().isBadRequest());
        mvc.perform(multipart(API).file(upload)
                        .param("removeEmptyColumns", "true").param("headerRows", "-1"))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Columns-Removed", "3"));
        mvc.perform(multipart(API).file(upload)
                        .param("removeEmptyColumns", "true").param("headerRows", "0"))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Columns-Removed", "2"));
    }

    @ParameterizedTest
    @ValueSource(strings = {"xlsx", "xls"})
    void compactsOnlyEmptyColumnsAndKeepsBodyValues(String extension) throws Exception {
        byte[] input;
        try (Workbook workbook = extension.equals("xls") ? new HSSFWorkbook() : new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Dispatch");
            var other = workbook.createSheet("Summary");
            sheet.createFreezePane(3, 4);
            Row header = sheet.createRow(3);
            String[] labels = {"Date", "Unused", "Stock", "Empty", " STATUS ", "Notes", "Tail", "Also empty"};
            for (int col = 0; col < labels.length; col++) {
                header.createCell(col).setCellValue(labels[col]);
            }
            var style = workbook.createCellStyle();
            style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            header.getCell(2).setCellStyle(style);
            sheet.setColumnWidth(2, 6200);
            sheet.setColumnHidden(2, true);
            sheet.addMergedRegion(new CellRangeAddress(3, 4, 2, 2));
            var conditional = sheet.getSheetConditionalFormatting();
            conditional.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(7, 7, 2, 2)},
                    conditional.createConditionalFormattingRule(ComparisonOperator.GT, "0"));
            var row = sheet.createRow(5);
            row.createCell(1).setCellValue(0);
            row.createCell(2).setCellValue(0);
            row.createCell(3).setCellValue(" NULL ");
            row.createCell(5).setCellFormula("C8+1");
            row.createCell(6).setCellValue(" 0 ");
            row.createCell(7).setCellValue("\u00a0");
            row.setHeight((short) 490);
            var finalRow = sheet.createRow(7);
            finalRow.createCell(2).setCellValue(24);
            finalRow.createCell(3).setBlank();
            var link = workbook.getCreationHelper().createHyperlink(org.apache.poi.common.usermodel.HyperlinkType.URL);
            link.setAddress("https://example.com/");
            finalRow.getCell(2).setHyperlink(link);
            other.createRow(0).createCell(0).setCellValue("Total");
            other.createRow(1).createCell(0).setCellFormula("Dispatch!C8");
            other.getRow(1).createCell(1).setCellFormula("Dispatch!H8");
            input = WorkbookTestData.bytes(workbook);
        }

        var response = mvc.perform(multipart(API).file(file("dispatch." + extension, input))
                        .param("removeEmptyColumns", "true"))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Columns-Removed", "4"))
                .andExpect(header().string("X-Cells-Cleared", "0"))
                .andReturn().getResponse();

        try (var result = WorkbookFactory.create(new ByteArrayInputStream(response.getContentAsByteArray()))) {
            var sheet = result.getSheet("Dispatch");
            assertThat(sheet.getPhysicalNumberOfRows()).isEqualTo(3);
            assertThat(sheet.getRow(6)).isNull();
            var header = sheet.getRow(3);
            assertThat(header.getLastCellNum()).isEqualTo((short) 4);
            assertThat(header.getCell(0).getStringCellValue()).isEqualTo("Date");
            assertThat(header.getCell(1).getStringCellValue()).isEqualTo("Stock");
            assertThat(header.getCell(2).getStringCellValue()).isEqualTo(" STATUS ");
            assertThat(header.getCell(3).getStringCellValue()).isEqualTo("Notes");
            assertThat(header.getCell(1).getCellStyle().getFillForegroundColor()).isEqualTo(IndexedColors.YELLOW.getIndex());
            assertThat(sheet.getColumnWidth(1)).isEqualTo(6200);
            assertThat(sheet.isColumnHidden(1)).isTrue();
            assertThat(sheet.isColumnHidden(2)).isFalse();
            assertThat(sheet.getMergedRegion(0).formatAsString()).isEqualTo("B4:B5");
            assertThat(sheet.getSheetConditionalFormatting().getConditionalFormattingAt(0)
                    .getFormattingRanges()[0].formatAsString()).isEqualTo("B8");
            assertThat(sheet.getPaneInformation().getVerticalSplitPosition()).isEqualTo((short) 2);
            assertThat(sheet.getRow(5).getHeight()).isEqualTo((short) 490);
            assertThat(sheet.getRow(5).getCell(1).getNumericCellValue()).isZero();
            assertThat(sheet.getRow(5).getCell(3).getCellFormula()).isEqualTo("B8+1");
            assertThat(sheet.getRow(7).getCell(1).getNumericCellValue()).isEqualTo(24);
            assertThat(sheet.getRow(7).getCell(1).getHyperlink().getAddress()).isEqualTo("https://example.com/");
            assertThat(result.getSheet("Summary").getRow(1).getCell(0).getCellFormula()).isEqualTo("Dispatch!B8");
            assertThat(result.getSheet("Summary").getRow(1).getCell(1).getCellFormula()).contains("#REF!");
        }
    }

    @Test
    void compactsAnXlsxAutofilterWithoutLosingColumnsWithOneValue() throws Exception {
        byte[] input;
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Truck Report");
            var header = sheet.createRow(0);
            String[] labels = {"Date", "Empty", "Cola", "Null", "Water", "Zero"};
            for (int col = 0; col < labels.length; col++) {
                header.createCell(col).setCellValue(labels[col]);
            }
            var first = sheet.createRow(1);
            first.createCell(1).setBlank();
            first.createCell(2).setCellValue(350);
            first.createCell(3).setCellValue(" NULL ");
            first.createCell(4).setCellValue(0);
            first.createCell(5).setCellValue("0");
            sheet.createRow(2).createCell(4).setCellValue(20);
            sheet.setAutoFilter(CellRangeAddress.valueOf("A1:F3"));
            var filter = sheet.getCTWorksheet().getAutoFilter();
            filter.addNewFilterColumn().setColId(1);
            filter.addNewFilterColumn().setColId(4);
            var sort = filter.addNewSortState();
            sort.setRef("A2:F3");
            sort.addNewSortCondition().setRef("E2:E3");
            input = WorkbookTestData.bytes(workbook);
        }
        var response = mvc.perform(multipart(API).file(file("truck.xlsx", input))
                        .param("removeEmptyColumns", "true"))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Columns-Removed", "3"))
                .andReturn().getResponse();
        try (var result = new XSSFWorkbook(new ByteArrayInputStream(response.getContentAsByteArray()))) {
            var sheet = result.getSheetAt(0);
            assertThat(sheet.getRow(0).getCell(1).getStringCellValue()).isEqualTo("Cola");
            assertThat(sheet.getRow(0).getCell(2).getStringCellValue()).isEqualTo("Water");
            assertThat(sheet.getRow(1).getCell(2).getNumericCellValue()).isZero();
            assertThat(sheet.getRow(2).getCell(2).getNumericCellValue()).isEqualTo(20);
            var filter = sheet.getCTWorksheet().getAutoFilter();
            assertThat(filter.getRef()).isEqualTo("A1:C3");
            assertThat(filter.sizeOfFilterColumnArray()).isEqualTo(1);
            assertThat(filter.getFilterColumnArray(0).getColId()).isEqualTo(2);
            assertThat(filter.getSortState().getSortConditionArray(0).getRef()).isEqualTo("C2:C3");
            assertThat(result.getNames("_xlnm._FilterDatabase")).singleElement()
                    .satisfies(name -> assertThat(name.getRefersToFormula()).isEqualTo("'Truck Report'!$A$1:$C$3"));
        }
    }

    @Test
    void keepsMultirowHeadersAndDropsAnUnusedMergedGroup() throws Exception {
        byte[] input;
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Inventory");
            sheet.createRow(0).createCell(0).setCellValue(" \u00a0");
            var group = sheet.createRow(1);
            group.createCell(0).setCellValue("Unused group");
            group.createCell(2).setCellValue("Available group");
            sheet.addMergedRegion(CellRangeAddress.valueOf("A2:B2"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("C2:E2"));
            var headers = sheet.createRow(2);
            for (int col = 0; col < 5; col++) {
                headers.createCell(col).setCellValue("Item " + col);
            }
            var values = sheet.createRow(3);
            values.createCell(0).setCellValue("0");
            values.createCell(1).setCellValue("\u00a0NULL\u00a0");
            values.createCell(3).setCellValue(20);
            values.createCell(4).setCellValue(0);
            var zeroRow = sheet.createRow(5);
            zeroRow.createCell(4).setCellValue(0);
            zeroRow.setHeight((short) 462);
            input = WorkbookTestData.bytes(workbook);
        }

        var response = mvc.perform(multipart(API).file(file("inventory.xlsx", input))
                        .param("removeEmptyColumns", "true").param("headerRows", "2"))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Columns-Removed", "3"))
                .andReturn().getResponse();
        try (var result = WorkbookFactory.create(new ByteArrayInputStream(response.getContentAsByteArray()))) {
            var sheet = result.getSheetAt(0);
            assertThat(sheet.getNumMergedRegions()).isEqualTo(1);
            assertThat(sheet.getMergedRegion(0).formatAsString()).isEqualTo("A2:B2");
            assertThat(sheet.getRow(1).getCell(0).getStringCellValue()).isEqualTo("Available group");
            assertThat(sheet.getRow(3).getCell(1).getNumericCellValue()).isEqualTo(20);
            assertThat(sheet.getRow(5).getPhysicalNumberOfCells()).isZero();
            assertThat(sheet.getRow(5).getHeight()).isEqualTo((short) 462);
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"xlsx", "xls"})
    void cleansAndPreservesLayout(String extension) throws Exception {
        var input = WorkbookTestData.sample(extension);
        var mimeType = extension.equals("xls") ? "application/vnd.ms-excel"
                : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        var response = mvc.perform(multipart(API).file(file("../report é." + extension, input)))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Worksheets-Processed", "2"))
                .andExpect(header().string("X-Cells-Cleared", "5"))
                .andExpect(header().exists("X-Processing-Millis"))
                .andExpect(header().string("Cache-Control", "no-store"))
                .andExpect(content().contentType(mimeType))
                .andReturn().getResponse();
        var disposition = ContentDisposition.parse(response.getHeader("Content-Disposition"));
        assertThat(disposition.getFilename()).isEqualTo("report é_cleaned." + extension);

        try (var cleaned = WorkbookFactory.create(new ByteArrayInputStream(response.getContentAsByteArray()))) {
            Sheet sheet = cleaned.getSheetAt(0);
            var row = sheet.getRow(1);
            for (int index : new int[]{0, 1, 2, 3, 7}) {
                assertThat(row.getCell(index).getCellType()).isEqualTo(CellType.BLANK);
            }
            assertThat(row.getCell(8)).isNull();
            assertThat(row.getCell(9).getCellType()).isEqualTo(CellType.ERROR);
            assertThat(row.getCell(0).getHyperlink().getAddress()).isEqualTo("https://example.com/");
            var style = row.getCell(0).getCellStyle();
            assertThat(style.getAlignment()).isEqualTo(HorizontalAlignment.CENTER);
            assertThat(style.getBorderBottom()).isEqualTo(BorderStyle.THIN);
            assertThat(style.getFillForegroundColor()).isEqualTo(IndexedColors.LIGHT_GREEN.getIndex());
            var font = cleaned.getFontAt(style.getFontIndex());
            assertThat(font.getColor()).isEqualTo(IndexedColors.BLUE.getIndex());
            assertThat(font.getBold()).isTrue();
            assertThat(style.getDataFormatString()).isEqualTo("0.000");
            assertThat(row.getCell(4).getCellFormula()).isEqualTo("1-1");
            assertThat(row.getCell(4).getNumericCellValue()).isZero();
            assertThat(row.getCell(5).getStringCellValue()).isEqualTo(" 0 ");
            assertThat(row.getCell(6).getBooleanCellValue()).isFalse();

            assertThat(sheet.getSheetName()).isEqualTo("Data 2026");
            assertThat(sheet.getMergedRegion(0).formatAsString()).isEqualTo("D4:E4");
            assertThat(sheet.getRow(3).getCell(3).getStringCellValue()).isEqualTo("Notes");
            assertThat(sheet.getPaneInformation().isFreezePane()).isTrue();
            assertThat(sheet.getPaneInformation().getHorizontalSplitPosition()).isEqualTo((short) 2);
            assertThat(sheet.getColumnWidth(3)).isEqualTo(6200);
            assertThat(sheet.isColumnHidden(1)).isTrue();
            assertThat(sheet.getRow(2).getZeroHeight()).isTrue();
            assertThat(row.getHeight()).isEqualTo((short) 520);
            assertThat(sheet.getSheetConditionalFormatting().getNumConditionalFormattings()).isEqualTo(1);
            assertThat(cleaned.getSheetAt(1).getSheetName()).isEqualTo("Archive");
            assertThat(cleaned.getSheetAt(1).getRow(3).getCell(4).getCellType()).isEqualTo(CellType.BLANK);
            assertThat(cleaned.isSheetHidden(1)).isTrue();
            if (cleaned instanceof HSSFWorkbook legacy) {
                assertThat(legacy.getSummaryInformation().getTitle()).isEqualTo("Fixture title");
            }
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"removeEmpty", "removeWhitespace", "removeNumericZero", "removeTextZero"})
    void rulesCanBeDisabledIndependently(String disabled) throws Exception {
        byte[] input = WorkbookTestData.sample("xlsx");
        var response = mvc.perform(multipart(API).file(file("report.xlsx", input)).param(disabled, "false"))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Cells-Cleared", disabled.equals("removeNumericZero") ? "3" : "4"))
                .andReturn().getResponse();

        try (var cleaned = WorkbookFactory.create(new ByteArrayInputStream(response.getContentAsByteArray()))) {
            var row = cleaned.getSheetAt(0).getRow(1);
            switch (disabled) {
                case "removeEmpty" -> assertThat(row.getCell(3).getStringCellValue()).isEmpty();
                case "removeWhitespace" -> assertThat(row.getCell(2).getStringCellValue()).isEqualTo(" \t\u00a0");
                case "removeNumericZero" -> assertThat(row.getCell(0).getCellType()).isEqualTo(CellType.NUMERIC);
                case "removeTextZero" -> assertThat(row.getCell(1).getStringCellValue()).isEqualTo("0");
            }
        }
    }

    @Test
    void returnsOriginalBytesWhenNoCellsChange() throws Exception {
        byte[] input = WorkbookTestData.sample("xlsx");
        mvc.perform(multipart(API).file(file("report.xlsx", input))
                        .param("removeEmpty", "false").param("removeWhitespace", "false")
                        .param("removeNumericZero", "false").param("removeTextZero", "false"))
                .andExpect(status().isOk())
                .andExpect(header().string("X-Cells-Cleared", "0"))
                .andExpect(content().bytes(input));
    }

    @Test
    void rejectsMissingEmptyOversizeAndInvalidUploads() throws Exception {
        mvc.perform(get("/api/excel/health")).andExpect(status().isOk()).andExpect(jsonPath("$.status").value("UP"));
        mvc.perform(multipart(API)).andExpect(status().isBadRequest()).andExpect(jsonPath("$.message").exists());
        mvc.perform(multipart(API).file(file("empty.xls", new byte[0]))).andExpect(status().isBadRequest());
        mvc.perform(multipart(API).file(file("large.xls", new byte[10 * 1024 * 1024 + 1])))
                .andExpect(status().isPayloadTooLarge());
        byte[] valid = WorkbookTestData.sample("xlsx");
        mvc.perform(multipart(API).file(file("book.csv", valid))).andExpect(status().isUnsupportedMediaType());
        mvc.perform(multipart(API).file(file("book.xls", valid))).andExpect(status().isUnsupportedMediaType());
        mvc.perform(multipart(API).file(file("book.xlsx", valid)).param("removeEmpty", "maybe"))
                .andExpect(status().isBadRequest());
        mvc.perform(multipart(API).file(file("fake.xlsx", "not a workbook".getBytes(StandardCharsets.UTF_8))))
                .andExpect(status().isUnprocessableEntity());
        mvc.perform(multipart(API).file(file("truncated.xlsx", Arrays.copyOf(valid, valid.length - 22))))
                .andExpect(status().isUnprocessableEntity());
        try (var empty = new XSSFWorkbook()) {
            mvc.perform(multipart(API).file(file("no-sheets.xlsx", WorkbookTestData.bytes(empty))))
                    .andExpect(status().isUnprocessableEntity())
                    .andExpect(jsonPath("$.message").value("The workbook does not contain any worksheets."));
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"a-test-password", Decryptor.DEFAULT_PASSWORD})
    void rejectsEncryptedXlsxEvenWithBuiltInPassword(String password) throws Exception {
        try (var fs = new POIFSFileSystem(); var output = new ByteArrayOutputStream()) {
            var encryptor = new EncryptionInfo(EncryptionMode.agile).getEncryptor();
            encryptor.confirmPassword(password);
            try (var stream = encryptor.getDataStream(fs)) {
                stream.write(WorkbookTestData.sample("xlsx"));
            }
            fs.writeFilesystem(output);
            assertEncrypted("encrypted.xlsx", output.toByteArray());
        }
    }

    @Test
    void rejectsEncryptedXls() throws Exception {
        try (var workbook = new HSSFWorkbook()) {
            workbook.createSheet("Protected").createRow(0).createCell(0).setCellValue(0);
            byte[] protectedBytes;
            try {
                Biff8EncryptionKey.setCurrentUserPassword("a-test-password");
                protectedBytes = WorkbookTestData.bytes(workbook);
            } finally {
                Biff8EncryptionKey.setCurrentUserPassword(null);
            }
            assertEncrypted("protected.xls", protectedBytes);
        }
    }

    private void assertEncrypted(String filename, byte[] bytes) throws Exception {
        mvc.perform(multipart(API).file(file(filename, bytes)))
                .andExpect(status().isUnprocessableEntity())
                .andExpect(jsonPath("$.message").value(org.hamcrest.Matchers.containsString("password protected")));
    }

    private MockMultipartFile file(String name, byte[] data) {
        return new MockMultipartFile("file", name, "application/octet-stream", data);
    }
}
