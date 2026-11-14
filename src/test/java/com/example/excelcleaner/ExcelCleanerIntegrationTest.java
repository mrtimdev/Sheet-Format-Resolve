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
