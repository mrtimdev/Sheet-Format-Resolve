package com.example.excelcleaner;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

final class WorkbookTestData {
    static byte[] sample(String ext) throws IOException {
        try (Workbook workbook = ext.equals("xls") ? new HSSFWorkbook() : new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data 2026");
            workbook.createSheet("Archive").createRow(3).createCell(4).setCellValue(0);
            workbook.setSheetHidden(1, true);
            sheet.createFreezePane(1, 2);
            sheet.setColumnHidden(1, true);
            sheet.setColumnWidth(3, 6200);
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 4));
            var conditional = sheet.getSheetConditionalFormatting();
            var rule = conditional.createConditionalFormattingRule(ComparisonOperator.GT, "0");
            rule.createPatternFormatting().setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
            conditional.addConditionalFormatting(new CellRangeAddress[]{new CellRangeAddress(1, 1, 0, 7)}, rule);

            Row row = sheet.createRow(1);
            row.setHeight((short) 520);
            var style = workbook.createCellStyle();
            var font = workbook.createFont();
            font.setBold(true);
            font.setColor(IndexedColors.BLUE.getIndex());
            style.setFont(font);
            style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            style.setBorderBottom(BorderStyle.THIN);
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setDataFormat(workbook.createDataFormat().getFormat("0.000"));
            Cell zero = row.createCell(0);
            zero.setCellStyle(style);
            zero.setCellValue(0);
            var link = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
            link.setAddress("https://example.com/");
            zero.setHyperlink(link);
            row.createCell(1).setCellValue("0");
            row.createCell(2).setCellValue(" \t\u00a0");
            row.createCell(3).setCellValue("");
            row.createCell(4).setCellFormula("1-1");
            row.createCell(5).setCellValue(" 0 ");
            row.createCell(6).setCellValue(false);
            row.createCell(7).setBlank();
            row.createCell(9).setCellErrorValue(FormulaError.NA.getCode());
            sheet.createRow(2).setZeroHeight(true);
            sheet.createRow(3).createCell(3).setCellValue("Notes");
            if (workbook instanceof HSSFWorkbook legacy) {
                legacy.createInformationProperties();
                legacy.getSummaryInformation().setTitle("Fixture title");
            }
            workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            return bytes(workbook);
        }
    }

    static byte[] bytes(Workbook workbook) throws IOException {
        try (var out = new ByteArrayOutputStream()) {
            workbook.write(out);
            return out.toByteArray();
        }
    }

    public static void main(String[] args) throws IOException {
        Files.write(Path.of(args[0]), sample("xlsx"));
    }
}
