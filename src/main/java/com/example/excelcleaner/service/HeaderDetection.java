package com.example.excelcleaner.service;

import java.util.HashSet;
import java.util.Locale;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

final class HeaderDetection {
    private static final Set<String> COMMON_LABELS = Set.of(
            "date", "id", "name", "status", "total", "remarks");

    private HeaderDetection() {}

    static int firstPopulatedRow(Sheet sheet) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (isPopulated(cell)) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    static int resolve(Sheet sheet, int first, int requested) {
        if (first < 0) {
            return 0;
        }
        if (requested >= 0) {
            return requested;
        }
        if (sheet instanceof XSSFSheet modern && modern.getCTWorksheet().isSetAutoFilter()) {
            int filterRow = CellRangeAddress.valueOf(modern.getCTWorksheet().getAutoFilter().getRef()).getFirstRow();
            if (filterRow >= first && filterRow - first < 100) {
                return filterRow - first + 1;
            }
        }

        // Skip merged text titles and look for text labels directly above typed data.
        int labelsIndex = first;
        var merges = sheet.getMergedRegions();
        while (labelsIndex - first < 7 && isTextRow(sheet.getRow(labelsIndex))) {
            int rowIndex = labelsIndex;
            var title = merges.stream().filter(merge -> merge.getFirstRow() == rowIndex
                    && merge.getLastColumn() > merge.getFirstColumn()).findFirst();
            if (title.isEmpty() || populatedCount(sheet.getRow(rowIndex)) != 1) {
                break;
            }
            labelsIndex = Math.min(first + 7, title.get().getLastRow() + 1);
        }
        Row labels = sheet.getRow(labelsIndex);
        if (!isTextRow(labels)) {
            return 0;
        }
        var recognized = new HashSet<String>();
        for (Cell cell : labels) {
            if (cell.getCellType() == CellType.STRING) {
                String value = cell.getStringCellValue().strip().toLowerCase(Locale.ROOT);
                if (COMMON_LABELS.contains(value)) {
                    recognized.add(value);
                }
            }
        }
        if (recognized.size() >= 2 && sheet.getLastRowNum() > labelsIndex) {
            return labelsIndex - first + 1;
        }
        for (int i = labelsIndex + 1; i <= Math.min(labelsIndex + 5, sheet.getLastRowNum()); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            for (Cell cell : row) {
                Cell label = labels.getCell(cell.getColumnIndex());
                if (isPopulated(label) && switch (cell.getCellType()) {
                    case BOOLEAN, NUMERIC, FORMULA -> true;
                    default -> false;
                }) {
                    return labelsIndex - first + 1;
                }
            }
            // A second text row can be data. Leave the decision to the user.
            if (isTextRow(row)) {
                break;
            }
        }
        return 0;
    }

    private static int populatedCount(Row row) {
        int count = 0;
        for (Cell cell : row) {
            if (isPopulated(cell)) {
                count++;
            }
        }
        return count;
    }

    private static boolean isTextRow(Row row) {
        if (row == null) {
            return false;
        }
        boolean foundText = false;
        for (Cell cell : row) {
            if (isPopulated(cell)) {
                if (cell.getCellType() != CellType.STRING) {
                    return false;
                }
                foundText = true;
            }
        }
        return foundText;
    }

    static boolean isPopulated(Cell cell) {
        return cell != null && switch (cell.getCellType()) {
            case BLANK, _NONE -> false;
            case STRING -> !cell.getStringCellValue().codePoints().allMatch(c ->
                    Character.isWhitespace(c) || Character.isSpaceChar(c) || c == 0x85);
            default -> true;
        };
    }
}
