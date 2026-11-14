package com.example.excelcleaner.model;

public record CleaningResult(String filename, String contentType, byte[] bytes,
                             int worksheets, long cellsCleared, long durationMillis) {
}
