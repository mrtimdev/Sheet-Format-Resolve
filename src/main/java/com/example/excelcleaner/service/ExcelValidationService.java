package com.example.excelcleaner.service;

import static com.example.excelcleaner.config.WebConfig.*;

import com.example.excelcleaner.exception.WorkbookException;
import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Locale;
import java.util.zip.CRC32;
import java.util.zip.ZipInputStream;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.apache.commons.compress.utils.SeekableInMemoryByteChannel;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelValidationService {
    public static final long MAX_UPLOAD_BYTES = 10L * 1024 * 1024;

    public enum Format {
        XLSX(".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        XLS(".xls", "application/vnd.ms-excel");

        private final String extension;
        private final String mediaType;

        Format(String extension, String mediaType) {
            this.extension = extension;
            this.mediaType = mediaType;
        }

        public String mediaType() {
            return mediaType;
        }
    }

    public Format validateUpload(MultipartFile file) {
        if (file.isEmpty()) {
            throw new WorkbookException(HttpStatus.BAD_REQUEST, "The uploaded file is empty.");
        }
        if (file.getSize() > MAX_UPLOAD_BYTES) {
            throw WorkbookException.tooLarge();
        }

        String filename = file.getOriginalFilename();
        if (filename != null) {
            for (Format format : Format.values()) {
                if (filename.toLowerCase(Locale.ROOT).endsWith(format.extension)) {
                    return format;
                }
            }
        }
        throw new WorkbookException(HttpStatus.UNSUPPORTED_MEDIA_TYPE, "Choose an .xlsx or .xls Excel workbook.");
    }

    public FileMagic validateContainer(byte[] bytes) throws IOException {
        FileMagic magic;
        try (var input = new BufferedInputStream(new ByteArrayInputStream(bytes))) {
            magic = FileMagic.valueOf(input);
        }

        switch (magic) {
            case OOXML -> validateArchive(bytes);
            case OLE2 -> { /* Includes encrypted OOXML; WorkbookFactory detects it. */ }
            default -> throw WorkbookException.invalid();
        }
        return magic;
    }

    private void validateArchive(byte[] bytes) throws IOException {
        // POI opens local ZIP records from an InputStream. Require a matching central
        // directory too, so truncated archives and ambiguous entries are rejected.
        try (var archive = ZipFile.builder()
                .setSeekableByteChannel(new SeekableInMemoryByteChannel(bytes)).get();
             var stream = new ZipInputStream(new ByteArrayInputStream(bytes))) {
            var names = new HashSet<String>();
            var buffer = new byte[8192];
            long expanded = 0;
            var entries = archive.getEntries();
            while (entries.hasMoreElements()) {
                var entry = entries.nextElement();
                if (!names.add(entry.getName().toLowerCase(Locale.ROOT))) {
                    throw WorkbookException.invalid();
                }
                if (names.size() > MAX_ENTRIES || entry.getSize() > MAX_ENTRY_BYTES ||
                        (expanded += entry.getSize()) > MAX_EXPANDED_BYTES) {
                    throw WorkbookException.tooLarge();
                }
            }
            int actualCount = 0;
            var localNames = new HashSet<String>();
            expanded = 0;
            for (var local = stream.getNextEntry(); local != null; local = stream.getNextEntry()) {
                var expected = archive.getEntry(local.getName());
                if (expected == null || !localNames.add(local.getName()) || ++actualCount > names.size()) {
                    throw WorkbookException.invalid();
                }
                var crc = new CRC32();
                long size = 0;
                for (int read; (read = stream.read(buffer)) != -1; ) {
                    if ((size += read) > MAX_ENTRY_BYTES || (expanded += read) > MAX_EXPANDED_BYTES) {
                        throw WorkbookException.tooLarge();
                    }
                    crc.update(buffer, 0, read);
                }
                if (size != expected.getSize() || crc.getValue() != expected.getCrc()) {
                    throw WorkbookException.invalid();
                }
            }
            if (actualCount == 0 || actualCount != names.size()) {
                throw WorkbookException.invalid();
            }
        }
    }

    public void validateWorkbook(Workbook workbook, Format format) {
        if (workbook instanceof HSSFWorkbook legacy && legacy.getEncryptionInfo() != null) {
            // POI can open some XLS files with Excel's built-in default password.
            throw WorkbookException.encrypted();
        }
        boolean matches = switch (format) {
            case XLSX -> workbook instanceof XSSFWorkbook modern
                    && modern.getWorkbookType() == XSSFWorkbookType.XLSX;
            case XLS -> workbook instanceof HSSFWorkbook;
        };
        if (!matches) {
            throw new WorkbookException(HttpStatus.UNSUPPORTED_MEDIA_TYPE,
                    "The workbook format does not match its extension. Save it as .xlsx or .xls in Excel and try again.");
        }
        if (workbook.getNumberOfSheets() == 0) {
            throw new WorkbookException(HttpStatus.UNPROCESSABLE_ENTITY, "The workbook does not contain any worksheets.");
        }
    }

    public String outputFilename(String original, Format format) {
        String name = original.replace('\\', '/');
        name = name.substring(name.lastIndexOf('/') + 1);
        String stem = name.length() >= format.extension.length()
                ? name.substring(0, name.length() - format.extension.length()) : "";
        stem = stem.replaceAll("[^\\p{L}\\p{N}\\p{M} ._()-]", "_").strip();
        if (stem.isBlank() || stem.chars().allMatch(c -> c == '.')) {
            stem = "workbook";
        }
        if (stem.codePointCount(0, stem.length()) > 96) {
            stem = stem.substring(0, stem.offsetByCodePoints(0, 96)).stripTrailing();
        }
        return stem + "_cleaned" + format.extension;
    }
}
