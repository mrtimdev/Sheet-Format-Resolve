package com.example.excelcleaner.controller;

import com.example.excelcleaner.exception.WorkbookException;
import com.example.excelcleaner.model.CleaningOptions;
import com.example.excelcleaner.model.WorkbookPreview;
import com.example.excelcleaner.service.ExcelCleanerService;
import java.nio.charset.StandardCharsets;
import java.util.Map;
import org.springframework.http.CacheControl;
import org.springframework.http.ContentDisposition;
import org.springframework.http.MediaType;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api/excel")
public class ExcelCleanerController {
    private final ExcelCleanerService cleaner;

    public ExcelCleanerController(ExcelCleanerService cleaner) {
        this.cleaner = cleaner;
    }

    @GetMapping("/health")
    public ResponseEntity<Map<String, String>> health() {
        return ResponseEntity.ok().cacheControl(CacheControl.noStore()).body(Map.of("status", "UP"));
    }

    @PostMapping(path = "/clean", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> clean(@RequestParam MultipartFile file,
                                        @RequestParam(defaultValue = "true") boolean removeEmpty,
                                        @RequestParam(defaultValue = "true") boolean removeWhitespace,
                                        @RequestParam(defaultValue = "true") boolean removeNumericZero,
                                        @RequestParam(defaultValue = "true") boolean removeTextZero,
                                        @RequestParam(defaultValue = "false") boolean removeEmptyColumns,
                                        @RequestParam(defaultValue = "1") int headerRows) {
        if (headerRows < -1 || headerRows > 100) {
            throw new WorkbookException(HttpStatus.BAD_REQUEST, "Header rows must be -1 for automatic detection, or 0 to 100.");
        }
        var options = new CleaningOptions(removeEmpty, removeWhitespace, removeNumericZero, removeTextZero,
                removeEmptyColumns, headerRows);
        var result = cleaner.clean(file, options);
        var disposition = ContentDisposition.attachment().filename(result.filename(), StandardCharsets.UTF_8).build();
        return ResponseEntity.ok()
                .cacheControl(CacheControl.noStore())
                .contentType(MediaType.parseMediaType(result.contentType()))
                .contentLength(result.bytes().length)
                .headers(headers -> headers.setContentDisposition(disposition))
                .header("X-Worksheets-Processed", String.valueOf(result.worksheets()))
                .header("X-Columns-Removed", String.valueOf(result.columnsRemoved()))
                .header("X-Cells-Cleared", String.valueOf(result.cellsCleared()))
                .header("X-Processing-Millis", String.valueOf(result.durationMillis()))
                .body(result.bytes());
    }

    @PostMapping(path = "/preview", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<WorkbookPreview> preview(@RequestParam MultipartFile file,
                                                    @RequestParam(defaultValue = "0") int sheetIndex,
                                                    @RequestParam(defaultValue = "0") int rowOffset,
                                                    @RequestParam(defaultValue = "0") int columnOffset) {
        return ResponseEntity.ok().cacheControl(CacheControl.noStore())
                .body(cleaner.preview(file, sheetIndex, rowOffset, columnOffset));
    }
}
