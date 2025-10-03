package com.usnbook.swagger2word.controller;

import com.fasterxml.jackson.databind.JsonNode;
import com.usnbook.swagger2word.service.ApiDocsService;
import com.usnbook.swagger2word.service.WordDocumentService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import reactor.core.publisher.Mono;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.MalformedURLException;
import java.net.URL;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@RestController
@RequestMapping("/api/generate-doc")
public class DocumentationController {

    private static final Logger logger = LoggerFactory.getLogger(DocumentationController.class);

    private final ApiDocsService apiDocsService;
    private final WordDocumentService wordDocumentService;

    public DocumentationController(ApiDocsService apiDocsService,
                                   WordDocumentService wordDocumentService) {
        this.apiDocsService = apiDocsService;
        this.wordDocumentService = wordDocumentService;
    }

    @GetMapping
    public Mono<ResponseEntity<byte[]>> generateDocumentation(@RequestParam String url) {
        if (url == null || url.trim().isEmpty()) {
            logger.warn("URL parameter is missing");
            return Mono.just(ResponseEntity.badRequest()
                    .contentType(MediaType.TEXT_PLAIN)
                    .body("Error: URL parameter is required".getBytes()));
        }

        try {
            new URL(url); // Валидация URL
        } catch (MalformedURLException e) {
            logger.warn("Invalid URL format: {}", url);
            return Mono.just(ResponseEntity.badRequest()
                    .contentType(MediaType.TEXT_PLAIN)
                    .body(("Error: Invalid URL format: " + url).getBytes()));
        }

        return apiDocsService.fetchApiDocs(url)
                .flatMap(apiSpec -> {
                    try {
                        logger.info("Generating Word document for API from URL: {}", url);

                        // Передаем JsonNode напрямую в сервис
                        byte[] content = wordDocumentService.generateWordDocumentFromJson(apiSpec);

                        String fileName = generateFileName(
                                apiSpec.path("info").path("title").asText("API_Documentation")
                        );

                        logger.info("Document generated successfully: {}", fileName);

                        return Mono.just(ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                        "attachment; filename=\"" + fileName + "\"")
                                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                                .contentLength(content.length)
                                .body(content));

                    } catch (Exception e) {
                        logger.error("Failed to generate document from URL: {}", url, e);
                        return Mono.error(new RuntimeException("Failed to generate document: " + e.getMessage(), e));
                    }
                })
                .onErrorResume(e -> {
                    logger.error("Error in documentation generation from URL: {}", url, e);
                    String errorMessage = e.getMessage() != null ? e.getMessage() : "Unknown error occurred";
                    return Mono.just(ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                            .contentType(MediaType.TEXT_PLAIN)
                            .body(("Error: " + errorMessage).getBytes()));
                });
    }

    private String generateFileName(String apiTitle) {
        if (apiTitle == null || apiTitle.trim().isEmpty()) {
            apiTitle = "API";
        }

        String safeTitle = apiTitle.replaceAll("[^a-zA-Z0-9а-яА-Я\\s-]", "_")
                .replaceAll("\\s+", "_")
                .substring(0, Math.min(50, apiTitle.length()));

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        return safeTitle + "_API_Documentation_" + timestamp + ".docx";
    }
}