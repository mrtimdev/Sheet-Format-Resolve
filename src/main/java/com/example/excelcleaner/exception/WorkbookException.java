package com.example.excelcleaner.exception;

import org.springframework.http.HttpStatus;

public class WorkbookException extends RuntimeException {
    private final HttpStatus status;

    public WorkbookException(HttpStatus status, String message) {
        super(message);
        this.status = status;
    }

    public HttpStatus status() {
        return status;
    }

    public static WorkbookException invalid() {
        return new WorkbookException(HttpStatus.UNPROCESSABLE_ENTITY,
                "This file is not a valid or supported Excel workbook. It may be corrupted or use an older Excel format.");
    }

    public static WorkbookException encrypted() {
        return new WorkbookException(HttpStatus.UNPROCESSABLE_ENTITY,
                "This workbook is password protected. Remove its opening password in Excel and try again.");
    }

    public static WorkbookException tooLarge() {
        return new WorkbookException(HttpStatus.PAYLOAD_TOO_LARGE,
                "This workbook exceeds the size or complexity limit. Choose a workbook up to 10 MB with less embedded data.");
    }
}
