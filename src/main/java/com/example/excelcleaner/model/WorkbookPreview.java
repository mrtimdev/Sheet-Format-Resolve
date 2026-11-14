package com.example.excelcleaner.model;

import java.util.List;

public record WorkbookPreview(List<String> sheetNames, int sheetIndex, int firstPopulatedRow,
                              int detectedHeaderRows, List<String> columns, List<PreviewRow> rows,
                              boolean moreColumns, int rowOffset, int columnOffset,
                              int totalRows, int totalColumns, boolean moreRows, int pageRows, int pageColumns) {
    public record PreviewRow(int rowNumber, List<String> values) {}
}
