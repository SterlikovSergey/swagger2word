package com.usnbook.swagger2word.controller;

import com.usnbook.swagger2word.service.ApiDocsService;
import com.usnbook.swagger2word.service.WordDocumentService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import reactor.core.publisher.Mono;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.nio.file.Files;

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
    public Mono<ResponseEntity<byte[]>> generateDocumentation() {
        return apiDocsService.fetchApiDocs()
                .map(apiSpec -> {
                    try {
                        logger.info("Generating Word document for API: {}", apiSpec.getInfo().getTitle());

                        String filePath = wordDocumentService.generateWordDocument(apiSpec);
                        File file = new File(filePath);

                        if (!file.exists()) {
                            throw new RuntimeException("Generated file not found: " + filePath);
                        }

                        byte[] content = Files.readAllBytes(file.toPath());

                        logger.info("Document generated successfully: {}", file.getName());

                        return ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                        "attachment; filename=\"" + file.getName() + "\"")
                                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                                .contentLength(content.length)
                                .body(content);

                    } catch (Exception e) {
                        logger.error("Failed to generate document", e);
                        throw new RuntimeException("Failed to generate document: " + e.getMessage(), e);
                    }
                })
                .onErrorResume(e -> {
                    logger.error("Error in documentation generation", e);
                    return Mono.just(ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                            .body(("Error: " + e.getMessage()).getBytes()));
                });
    }
}