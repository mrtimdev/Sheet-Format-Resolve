package com.example.excelcleaner.config;

import jakarta.annotation.PostConstruct;
import org.apache.coyote.http11.AbstractHttp11Protocol;
import org.apache.poi.openxml4j.util.ZipInputStreamZipEntrySource;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.boot.web.embedded.tomcat.TomcatServletWebServerFactory;
import org.springframework.boot.web.server.WebServerFactoryCustomizer;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class WebConfig {
    public static final int MAX_EXPANDED_BYTES = 48 * 1024 * 1024;
    public static final int MAX_ENTRY_BYTES = 20 * 1024 * 1024;
    public static final int MAX_ENTRIES = 1024;

    @PostConstruct
    void configurePoiLimits() {
        // A negative threshold keeps POI's stream-backed ZIP entries in memory.
        ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(-1);
        ZipSecureFile.setMaxEntrySize(MAX_ENTRY_BYTES);
        ZipSecureFile.setMaxFileCount(MAX_ENTRIES);
    }

    @Bean
    WebServerFactoryCustomizer<TomcatServletWebServerFactory> uploadTimeout() {
        return factory -> factory.addConnectorCustomizers(connector -> {
            if (connector.getProtocolHandler() instanceof AbstractHttp11Protocol<?> protocol) {
                protocol.setDisableUploadTimeout(false);
                protocol.setConnectionUploadTimeout(30_000);
            }
        });
    }
}
