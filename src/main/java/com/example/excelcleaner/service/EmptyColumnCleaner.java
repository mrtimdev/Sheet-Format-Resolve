package com.example.excelcleaner.service;

import com.example.excelcleaner.exception.WorkbookException;
import java.util.ArrayList;
import java.util.BitSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.stream.IntStream;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.record.AutoFilterInfoRecord;
import org.apache.poi.hssf.record.ColumnInfoRecord;
import org.apache.poi.hssf.record.aggregates.ColumnInfoRecordsAggregate;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.SheetNameFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.helpers.XSSFColumnShifter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;

@Service
public class EmptyColumnCleaner {
    private static final Set<String> PROTECTED_HEADERS = Set.of(
            "date", "id", "name", "status", "total", "remarks");

    public int removeEmptyColumns(Sheet sheet, int first, int headerRows) {
        if (first < 0) {
            return 0;
        }
        int headerEnd = first + headerRows;
        var keep = new BitSet();
        int start = Integer.MAX_VALUE;
        int end = -1;
        boolean hasBody = false;
        for (Row row : sheet) {
            if (row.getRowNum() >= headerEnd) {
                hasBody |= row.getPhysicalNumberOfCells() > 0;
            }
            for (Cell cell : row) {
                int col = cell.getColumnIndex();
                start = Math.min(start, col);
                end = Math.max(end, col);
                if (cell.getCellType() == CellType.FORMULA
                        || (row.getRowNum() < headerEnd ? protectedHeader(cell) : hasValue(cell))) {
                    keep.set(col);
                }
            }
        }
        if (!hasBody || end < start) {
            return 0;
        }

        var links = sheet.getHyperlinkList();
        for (var link : links) {
            keep.set(link.getFirstColumn(), link.getLastColumn() + 1);
        }
        for (var address : sheet.getCellComments().keySet()) {
            keep.set(address.getColumn());
        }
        List<CellRangeAddress> merges = sheet.getMergedRegions();
        for (var merge : merges) {
            int retained = keep.nextSetBit(merge.getFirstColumn());
            if (retained == -1 || retained > merge.getLastColumn()) {
                continue;
            }
            Row row = sheet.getRow(merge.getFirstRow());
            Cell anchor = row == null ? null : row.getCell(merge.getFirstColumn());
            // Keep a visible merge caption in place while its span is compacted.
            if (anchor != null && hasValue(anchor)) {
                keep.set(merge.getFirstColumn());
            }
        }

        var removed = new BitSet();
        removed.set(start, end + 1);
        removed.andNot(keep);
        if (removed.isEmpty()) {
            return 0;
        }
        ensureSupported(sheet);
        int[] offsets = new int[sheet.getWorkbook().getSpreadsheetVersion().getMaxColumns() + 1];
        for (int col = 0; col < offsets.length - 1; col++) {
            offsets[col + 1] = offsets[col] + (removed.get(col) ? 1 : 0);
        }

        var pane = sheet.getPaneInformation();
        int[] breaks = sheet.getColumnBreaks();
        var legacyLinks = sheet instanceof HSSFSheet ? links : List.<Hyperlink>of();
        sheet.removeMergedRegions(IntStream.range(0, merges.size()).boxed().toList());
        removeCells(sheet, removed);

        // Include coordinates used by hyperlinks and comments even when those
        // coordinates do not have cell records.
        int last = Math.max(end, keep.length() - 1);
        // Shift one contiguous removal at a time. Right-to-left keeps the original
        // indexes of the remaining removal blocks valid.
        for (int right = removed.length() - 1; right >= 0; ) {
            int left = removed.previousClearBit(right) + 1;
            int count = right - left + 1;
            int sourceEnd = Math.min(offsets.length - 2, Math.max(last, right + count));
            if (right < last) {
                sheet.shiftColumns(right + 1, sourceEnd, -count);
            } else if (right < sourceEnd) {
                // The rightmost block has no stored cells to move. Formula
                // references and formatting rules still need the same shift.
                updateReferences(sheet, right + 1, sourceEnd, -count);
            }
            last -= count;
            right = removed.previousSetBit(left - 1);
        }
        sheet.getWorkbook().setForceFormulaRecalculation(true);

        for (var link : legacyLinks) {
            int left = link.getFirstColumn();
            int right = link.getLastColumn();
            link.setFirstColumn(left - offsets[left]);
            link.setLastColumn(right - offsets[right + 1]);
        }
        for (var merge : merges) {
            var range = compactRange(merge, offsets);
            if (range != null && range.getNumberOfCells() > 1) {
                // The original regions were already valid. A monotonic projection
                // cannot introduce an overlap between surviving cells.
                sheet.addMergedRegionUnsafe(range);
            }
        }
        for (int col : breaks) {
            sheet.removeColumnBreak(col);
        }
        for (int col : breaks) {
            if (!removed.get(col)) {
                sheet.setColumnBreak(col - offsets[col]);
            }
        }
        if (pane != null && pane.isFreezePane()) {
            int split = pane.getVerticalSplitPosition();
            int visible = pane.getVerticalSplitLeftColumn();
            sheet.createFreezePane(split - offsets[split], pane.getHorizontalSplitPosition(),
                    visible - offsets[visible], pane.getHorizontalSplitTopRow());
        }
        compactColumnSettings(sheet, offsets);
        if (sheet instanceof XSSFSheet modern) {
            compactAutoFilter(modern, offsets);
        }
        return removed.cardinality();
    }

    private void updateReferences(Sheet sheet, int first, int last, int offset) {
        var workbook = sheet.getWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheet);
        if (sheet instanceof XSSFSheet modern) {
            var shift = FormulaShifter.createForColumnShift(sheetIndex, sheet.getSheetName(),
                    first, last, offset, workbook.getSpreadsheetVersion());
            var updater = new XSSFColumnShifter(modern);
            updater.updateFormulas(shift);
            updater.updateConditionalFormatting(shift);
            updater.updateNamedRanges(shift);
        } else {
            var legacy = (HSSFSheet) sheet;
            var records = legacy.getWorkbook().getWorkbook();
            var shift = FormulaShifter.createForColumnShift(records.checkExternSheet(sheetIndex),
                    sheet.getSheetName(), first, last, offset, workbook.getSpreadsheetVersion());
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                var target = (HSSFSheet) workbook.getSheetAt(i);
                target.getSheet().updateFormulasAfterCellShift(shift, records.checkExternSheet(i));
            }
            records.updateNamesAfterCellShift(shift);
        }
    }

    private boolean protectedHeader(Cell cell) {
        return cell.getCellType() == CellType.STRING
                && PROTECTED_HEADERS.contains(stripWhitespace(cell.getStringCellValue()).toLowerCase(Locale.ROOT));
    }

    private boolean hasValue(Cell cell) {
        return switch (cell.getCellType()) {
            case NUMERIC -> cell.getNumericCellValue() != 0.0;
            case STRING -> {
                String value = stripWhitespace(cell.getStringCellValue());
                yield !value.isEmpty() && !value.equals("0") && !value.equalsIgnoreCase("null");
            }
            case BLANK, _NONE -> false;
            default -> true;
        };
    }

    private String stripWhitespace(String value) {
        int start = 0;
        int end = value.length();
        while (start < end && isSpace(value.charAt(start))) {
            start++;
        }
        while (end > start && isSpace(value.charAt(end - 1))) {
            end--;
        }
        return value.substring(start, end);
    }

    private boolean isSpace(char value) {
        return Character.isSpaceChar(value) || Character.isWhitespace(value) || value == 0x85;
    }

    private void ensureSupported(Sheet sheet) {
        boolean unsupported = !sheet.getDataValidations().isEmpty();
        if (sheet instanceof XSSFSheet modern) {
            var xml = modern.getCTWorksheet();
            unsupported |= !modern.getTables().isEmpty() || !modern.getPivotTables().isEmpty()
                    || xml.isSetDrawing();
        } else if (sheet instanceof HSSFSheet legacy) {
            var drawing = legacy.getDrawingPatriarch();
            unsupported |= (drawing != null && !drawing.getChildren().isEmpty())
                    || legacy.getSheet().findFirstRecordBySid(AutoFilterInfoRecord.sid) != null;
        }
        for (Row row : sheet) {
            for (Cell cell : row) {
                unsupported |= cell.isPartOfArrayFormulaGroup();
            }
        }
        if (unsupported) {
            throw new WorkbookException(HttpStatus.UNPROCESSABLE_ENTITY,
                    "Empty columns cannot be safely removed from this sheet because it contains a table, legacy filter, "
                            + "drawing, data validation or array formula. Use Clear matching cells for this workbook.");
        }
    }

    private void removeCells(Sheet sheet, BitSet removed) {
        for (Row row : sheet) {
            var discard = new ArrayList<Cell>();
            for (Cell cell : row) {
                if (removed.get(cell.getColumnIndex())) {
                    discard.add(cell);
                }
            }
            discard.forEach(row::removeCell);
        }
    }

    private CellRangeAddress compactRange(CellRangeAddress range, int[] offsets) {
        int left = range.getFirstColumn() - offsets[range.getFirstColumn()];
        int right = range.getLastColumn() - offsets[range.getLastColumn() + 1];
        return right < left ? null : new CellRangeAddress(range.getFirstRow(), range.getLastRow(), left, right);
    }

    private void compactAutoFilter(XSSFSheet sheet, int[] offsets) {
        var worksheet = sheet.getCTWorksheet();
        if (!worksheet.isSetAutoFilter()) {
            return;
        }
        var filter = worksheet.getAutoFilter();
        var original = CellRangeAddress.valueOf(filter.getRef());
        var range = compactRange(original, offsets);
        var workbook = sheet.getWorkbook();
        int sheetIndex = workbook.getSheetIndex(sheet);
        var filterName = workbook.getNames("_xlnm._FilterDatabase").stream()
                .filter(name -> name.getSheetIndex() == sheetIndex).findFirst();
        if (range == null) {
            worksheet.unsetAutoFilter();
            filterName.ifPresent(workbook::removeName);
            return;
        }
        filter.setRef(range.formatAsString());
        filterName.ifPresent(name -> {
            var first = new CellReference(range.getFirstRow(), range.getFirstColumn(), true, true);
            var last = new CellReference(range.getLastRow(), range.getLastColumn(), true, true);
            name.setRefersToFormula(SheetNameFormatter.format(sheet.getSheetName()) + "!"
                    + first.formatAsString() + ":" + last.formatAsString());
        });
        for (int i = filter.sizeOfFilterColumnArray() - 1; i >= 0; i--) {
            var criterion = filter.getFilterColumnArray(i);
            int col = original.getFirstColumn() + Math.toIntExact(criterion.getColId());
            if (offsets[col] != offsets[col + 1]) {
                filter.removeFilterColumn(i);
            } else {
                criterion.setColId(col - offsets[col] - range.getFirstColumn());
            }
        }
        if (!filter.isSetSortState()) {
            return;
        }
        var sort = filter.getSortState();
        var sortRange = compactRange(CellRangeAddress.valueOf(sort.getRef()), offsets);
        if (sortRange == null) {
            filter.unsetSortState();
            return;
        }
        sort.setRef(sortRange.formatAsString());
        for (int i = sort.sizeOfSortConditionArray() - 1; i >= 0; i--) {
            var condition = sort.getSortConditionArray(i);
            var conditionRange = compactRange(CellRangeAddress.valueOf(condition.getRef()), offsets);
            if (conditionRange == null) {
                sort.removeSortCondition(i);
            } else {
                condition.setRef(conditionRange.formatAsString());
            }
        }
        if (sort.sizeOfSortConditionArray() == 0) {
            filter.unsetSortState();
        }
    }

    private void compactColumnSettings(Sheet sheet, int[] offsets) {
        if (sheet instanceof XSSFSheet modern) {
            var worksheet = modern.getCTWorksheet();
            for (int i = worksheet.sizeOfColsArray() - 1; i >= 0; i--) {
                var group = worksheet.getColsArray(i);
                var columns = new ArrayList<CTCol>();
                for (var original : group.getColList()) {
                    int left = Math.toIntExact(original.getMin()) - 1;
                    int right = Math.toIntExact(original.getMax()) - 1;
                    int first = left - offsets[left];
                    int last = right - offsets[right + 1];
                    if (last >= first) {
                        var column = (CTCol) original.copy();
                        column.setMin(first + 1);
                        column.setMax(last + 1);
                        columns.add(column);
                    }
                }
                if (columns.isEmpty()) {
                    worksheet.removeCols(i);
                } else {
                    group.setColArray(columns.toArray(CTCol[]::new));
                }
            }
            return;
        }
        // Replace the serialized BIFF column records after all public sheet edits.
        var records = ((HSSFSheet) sheet).getSheet().getRecords();
        for (int i = 0; i < records.size(); i++) {
            if (records.get(i) instanceof ColumnInfoRecordsAggregate original) {
                var compacted = new ColumnInfoRecordsAggregate();
                original.visitContainedRecords(record -> {
                    var column = (ColumnInfoRecord) record;
                    int left = column.getFirstColumn();
                    int right = column.getLastColumn();
                    int first = left - offsets[left];
                    int last = right - offsets[right + 1];
                    if (last >= first) {
                        var copy = column.copy();
                        copy.setFirstColumn(first);
                        copy.setLastColumn(last);
                        compacted.insertColumn(copy);
                    }
                });
                records.set(i, compacted);
                return;
            }
        }
    }
}
