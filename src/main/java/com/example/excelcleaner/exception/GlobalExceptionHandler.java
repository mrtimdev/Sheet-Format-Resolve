package com.example.excelcleaner.exception;

import java.util.Map;
import org.springframework.http.CacheControl;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;
import org.springframework.web.multipart.MaxUploadSizeExceededException;
import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.support.MissingServletRequestPartException;
import org.springframework.web.method.annotation.MethodArgumentTypeMismatchException;

@RestControllerAdvice
public class GlobalExceptionHandler {
    @ExceptionHandler(WorkbookException.class)
    ResponseEntity<Map<String, String>> workbook(WorkbookException error) {
        return response(error.status(), error.getMessage());
    }

    @ExceptionHandler(MaxUploadSizeExceededException.class)
    ResponseEntity<Map<String, String>> size() {
        return workbook(WorkbookException.tooLarge());
    }

    @ExceptionHandler(MissingServletRequestPartException.class)
    ResponseEntity<Map<String, String>> missing() {
        return response(HttpStatus.BAD_REQUEST, "Choose an Excel workbook to upload.");
    }

    @ExceptionHandler(MultipartException.class)
    ResponseEntity<Map<String, String>> multipart() {
        return response(HttpStatus.BAD_REQUEST, "The upload could not be read. Send one workbook as multipart form data.");
    }

    @ExceptionHandler(MethodArgumentTypeMismatchException.class)
    ResponseEntity<Map<String, String>> option() {
        return response(HttpStatus.BAD_REQUEST, "Cleaning options must be true or false.");
    }

    private ResponseEntity<Map<String, String>> response(HttpStatus status, String message) {
        return ResponseEntity.status(status).cacheControl(CacheControl.noStore()).body(Map.of("message", message));
    }
}
