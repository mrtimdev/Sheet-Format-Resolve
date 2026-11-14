package com.example.excelcleaner.model;

public record CleaningOptions(boolean removeEmpty, boolean removeWhitespace,
                              boolean removeNumericZero, boolean removeTextZero) {
}
