package com.example.excelcleaner.service;

import com.example.excelcleaner.config.WebConfig;
import com.example.excelcleaner.exception.WorkbookException;
import com.example.excelcleaner.model.CleaningOptions;
import com.example.excelcleaner.model.CleaningResult;
import com.example.excelcleaner.model.WorkbookPreview;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.concurrent.Semaphore;
import java.util.concurrent.TimeUnit;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelCleanerService {
    private static final int PREVIEW_ROWS = 30;
    private static final int PREVIEW_COLUMNS = 20;
    private final ExcelValidationService validation;
    private final EmptyColumnCleaner columnCleaner;
    private final Semaphore processingSlots = new Semaphore(2);

    public ExcelCleanerService(ExcelValidationService validation, EmptyColumnCleaner columnCleaner) {
        this.validation = validation;
        this.columnCleaner = columnCleaner;
    }

    public CleaningResult clean(MultipartFile file, CleaningOptions options) {
        var format = validation.validateUpload(file);
        acquireProcessingSlot();
        try {
            return clean(file, format, options);
        } finally {
            processingSlots.release();
        }
    }

    private void acquireProcessingSlot() {
        if (!processingSlots.tryAcquire()) {
            throw new WorkbookException(HttpStatus.SERVICE_UNAVAILABLE,
                    "The cleaner is busy. Try again in a moment.");
        }
    }

    private byte[] readUpload(MultipartFile file) {
        try {
            return file.getBytes();
        } catch (IOException error) {
            throw new WorkbookException(HttpStatus.BAD_REQUEST, "The uploaded file could not be read.");
        }
    }

    private CleaningResult clean(MultipartFile file, ExcelValidationService.Format format, CleaningOptions options) {
        long start = System.nanoTime();
        byte[] input = readUpload(file);
        try (var workbook = open(input)) {
            validation.validateWorkbook(workbook, format);
            long cleared = 0;
            int columnsRemoved = 0;
            for (var sheet : workbook) {
                if (options.removeEmptyColumns()) {
                    int first = HeaderDetection.firstPopulatedRow(sheet);
                    int headerRows = HeaderDetection.resolve(sheet, first, options.headerRows());
                    columnsRemoved += columnCleaner.removeEmptyColumns(sheet, first, headerRows);
                    continue;
                }
                for (var row : sheet) {
                    for (var cell : row) {
                        if (shouldClear(cell, options)) {
                            cell.setBlank();
                            cleared++;
                        }
                    }
                }
            }
            var output = cleared == 0 && columnsRemoved == 0 ? input : write(workbook);
            return new CleaningResult(validation.outputFilename(file.getOriginalFilename(), format),
                    format.mediaType(), output, workbook.getNumberOfSheets(), cleared, columnsRemoved,
                    TimeUnit.NANOSECONDS.toMillis(System.nanoTime() - start));
        } catch (WorkbookException error) {
            throw error;
        } catch (IOException | RuntimeException error) {
            throw new WorkbookException(HttpStatus.INTERNAL_SERVER_ERROR,
                    "The workbook could not be cleaned safely. Try saving a fresh copy in Excel.");
        }
    }

    public WorkbookPreview preview(MultipartFile file, int sheetIndex, int rowOffset, int columnOffset) {
        if (rowOffset < 0 || columnOffset < 0) {
            throw invalidPreviewPage();
        }
        var format = validation.validateUpload(file);
        acquireProcessingSlot();
        try (var workbook = open(readUpload(file))) {
            validation.validateWorkbook(workbook, format);
            if (sheetIndex < 0 || sheetIndex >= workbook.getNumberOfSheets()) {
                throw new WorkbookException(HttpStatus.BAD_REQUEST, "Choose a worksheet in this workbook.");
            }
            var names = new ArrayList<String>();
            for (var sheet : workbook) {
                names.add(sheet.getSheetName());
            }
            var sheet = workbook.getSheetAt(sheetIndex);
            var columns = new ArrayList<String>();
            var rows = new ArrayList<WorkbookPreview.PreviewRow>();
            int firstRow = -1;
            int lastRow = -1;
            int firstColumn = workbook.getSpreadsheetVersion().getMaxColumns();
            int lastColumn = -1;
            for (var row : sheet) {
                for (var cell : row) {
                    if (HeaderDetection.isPopulated(cell)) {
                        if (firstRow < 0) {
                            firstRow = row.getRowNum();
                        }
                        lastRow = row.getRowNum();
                        firstColumn = Math.min(firstColumn, cell.getColumnIndex());
                        lastColumn = Math.max(lastColumn, cell.getColumnIndex());
                    }
                }
            }
            if (firstRow < 0 && rowOffset == 0 && columnOffset == 0) {
                return new WorkbookPreview(names, sheetIndex, 0, 0, columns, rows,
                        false, 0, 0, 0, 0, false, PREVIEW_ROWS, PREVIEW_COLUMNS);
            }
            int totalRows = lastRow - firstRow + 1;
            int totalColumns = lastColumn - firstColumn + 1;
            if (firstRow < 0 || rowOffset >= totalRows || columnOffset >= totalColumns) {
                throw invalidPreviewPage();
            }
            int detected = HeaderDetection.resolve(sheet, firstRow, -1);
            int rowStart = firstRow + rowOffset;
            int rowEnd = Math.min(lastRow, rowStart + PREVIEW_ROWS - 1);
            int columnStart = firstColumn + columnOffset;
            int columnEnd = Math.min(lastColumn, columnStart + PREVIEW_COLUMNS - 1);
            for (int col = columnStart; col <= columnEnd; col++) {
                columns.add(CellReference.convertNumToColString(col));
            }
            var formatter = new DataFormatter();
            formatter.setUseCachedValuesForFormulaCells(true);
            for (int i = rowStart; i <= rowEnd; i++) {
                var row = sheet.getRow(i);
                var values = new ArrayList<String>();
                for (int col = columnStart; col <= columnEnd; col++) {
                    values.add(formatter.formatCellValue(row == null ? null : row.getCell(col)));
                }
                rows.add(new WorkbookPreview.PreviewRow(i + 1, values));
            }
            return new WorkbookPreview(names, sheetIndex, firstRow + 1, detected, columns, rows,
                    columnEnd < lastColumn, rowOffset, columnOffset, totalRows, totalColumns, rowEnd < lastRow,
                    PREVIEW_ROWS, PREVIEW_COLUMNS);
        } catch (WorkbookException error) {
            throw error;
        } catch (IOException | RuntimeException error) {
            throw WorkbookException.invalid();
        } finally {
            processingSlots.release();
        }
    }

    private WorkbookException invalidPreviewPage() {
        return new WorkbookException(HttpStatus.BAD_REQUEST, "Choose rows and columns within this worksheet.");
    }

    private Workbook open(byte[] bytes) {
        try {
            var magic = validation.validateContainer(bytes);
            // WorkbookFactory normally tries Excel's built-in password on protected
            // OOXML. Check the wrapper first to enforce rejection of all encryption.
            if (magic == FileMagic.OLE2) {
                var compound = new POIFSFileSystem(new ByteArrayInputStream(bytes));
                boolean transferred = false;
                try {
                    if (compound.getRoot().hasEntryCaseInsensitive("EncryptedPackage")) {
                        throw WorkbookException.encrypted();
                    }
                    var workbook = WorkbookFactory.create(compound.getRoot(), null);
                    // HSSFWorkbook owns the OLE2 filesystem, including preserved records.
                    transferred = true;
                    return workbook;
                } finally {
                    if (!transferred) {
                        compound.close();
                    }
                }
            }
            return WorkbookFactory.create(new ByteArrayInputStream(bytes));
        } catch (EncryptedDocumentException error) {
            throw WorkbookException.encrypted();
        } catch (WorkbookException error) {
            throw error;
        } catch (IOException | RuntimeException error) {
            throw WorkbookException.invalid();
        }
    }

    private boolean shouldClear(Cell cell, CleaningOptions options) {
        return switch (cell.getCellType()) {
            case NUMERIC -> options.removeNumericZero() && cell.getNumericCellValue() == 0.0;
            case STRING -> {
                String value = cell.getStringCellValue();
                if (value == null || value.isEmpty()) {
                    yield options.removeEmpty();
                }
                yield (options.removeTextZero() && value.equals("0"))
                        || (options.removeWhitespace()
                        && value.codePoints().allMatch(c -> Character.isWhitespace(c) || Character.isSpaceChar(c) || c == 0x85));
            }
            default -> false; // Formulas, errors, booleans and already blank cells.
        };
    }

    private byte[] write(Workbook workbook) throws IOException {
        var bytes = new ByteArrayOutputStream();
        // Cap the serialized workbook too, independent of its uploaded size.
        var bounded = new OutputStream() {
            @Override
            public void write(int value) {
                check(1);
                bytes.write(value);
            }

            @Override
            public void write(byte[] data, int offset, int length) {
                check(length);
                bytes.write(data, offset, length);
            }

            private void check(int length) {
                if (length > WebConfig.MAX_EXPANDED_BYTES - bytes.size()) {
                    throw WorkbookException.tooLarge();
                }
            }
        };
        workbook.write(bounded);
        return bytes.toByteArray();
    }
}
