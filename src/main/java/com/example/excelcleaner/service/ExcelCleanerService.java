package com.example.excelcleaner.service;

import com.example.excelcleaner.config.WebConfig;
import com.example.excelcleaner.exception.WorkbookException;
import com.example.excelcleaner.model.CleaningOptions;
import com.example.excelcleaner.model.CleaningResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.concurrent.Semaphore;
import java.util.concurrent.TimeUnit;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelCleanerService {
    private final ExcelValidationService validation;
    private final Semaphore processingSlots = new Semaphore(2);

    public ExcelCleanerService(ExcelValidationService validation) {
        this.validation = validation;
    }

    public CleaningResult clean(MultipartFile file, CleaningOptions options) {
        var format = validation.validateUpload(file);
        if (!processingSlots.tryAcquire()) {
            throw new WorkbookException(HttpStatus.SERVICE_UNAVAILABLE,
                    "The cleaner is busy. Try again in a moment.");
        }
        try {
            return clean(file, format, options);
        } finally {
            processingSlots.release();
        }
    }

    private CleaningResult clean(MultipartFile file, ExcelValidationService.Format format, CleaningOptions options) {
        long start = System.nanoTime();
        byte[] input;
        try {
            input = file.getBytes();
        } catch (IOException error) {
            throw new WorkbookException(HttpStatus.BAD_REQUEST, "The uploaded file could not be read.");
        }

        try (var workbook = open(input)) {
            validation.validateWorkbook(workbook, format);
            long cleared = 0;
            for (var sheet : workbook) {
                for (var row : sheet) {
                    for (var cell : row) {
                        if (shouldClear(cell, options)) {
                            cell.setBlank();
                            cleared++;
                        }
                    }
                }
            }
            var output = cleared == 0 ? input : write(workbook);
            return new CleaningResult(validation.outputFilename(file.getOriginalFilename(), format),
                    format.mediaType(), output, workbook.getNumberOfSheets(), cleared,
                    TimeUnit.NANOSECONDS.toMillis(System.nanoTime() - start));
        } catch (WorkbookException error) {
            throw error;
        } catch (IOException | RuntimeException error) {
            throw new WorkbookException(HttpStatus.INTERNAL_SERVER_ERROR,
                    "The workbook could not be cleaned safely. Try saving a fresh copy in Excel.");
        }
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
