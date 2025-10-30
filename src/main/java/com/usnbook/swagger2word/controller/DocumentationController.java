package com.usnbook.swagger2word.controller;

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

import java.io.File;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Files;
import java.time.Duration;

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
    public Mono<ResponseEntity<byte[]>> generateDocumentation(@RequestParam(required = false) String url) {
        // Валидация входного параметра
        if (url == null || url.trim().isEmpty()) {
            logger.warn("URL parameter is missing");
            return Mono.just(ResponseEntity.badRequest()
                    .contentType(MediaType.TEXT_PLAIN)
                    .body("Error: URL parameter is required. Example: ?url=https://api.example.com/docs".getBytes()));
        }

        String cleanedUrl = url.trim();

        // Валидация формата URL
        try {
            URL parsedUrl = new URL(cleanedUrl);
            if (!parsedUrl.getProtocol().startsWith("http")) {
                throw new MalformedURLException("Only HTTP/HTTPS protocols are supported");
            }
        } catch (MalformedURLException e) {
            logger.warn("Invalid URL format: {}", cleanedUrl);
            return Mono.just(ResponseEntity.badRequest()
                    .contentType(MediaType.TEXT_PLAIN)
                    .body(("Error: Invalid URL format: " + cleanedUrl + ". " + e.getMessage()).getBytes()));
        }

        logger.info("Starting documentation generation for URL: {}", cleanedUrl);

        return apiDocsService.fetchApiDocs(cleanedUrl)
                .flatMap(apiSpec -> {
                    try {
                        // Проверка валидности полученной спецификации
                        if (apiSpec == null) {
                            throw new RuntimeException("Received null API specification from the server");
                        }

                        if (apiSpec.getInfo() == null) {
                            throw new RuntimeException("API specification does not contain required 'info' section");
                        }

                        String apiTitle = apiSpec.getInfo().getTitle();
                        logger.info("Generating Word document for API: {} from URL: {}", apiTitle, cleanedUrl);

                        //  создаем временный файл
                        String filePath = wordDocumentService.generateWordDocument(apiSpec);
                        File file = new File(filePath);

                        // Проверка существования файла
                        if (!file.exists()) {
                            throw new RuntimeException("Generated file not found: " + filePath);
                        }

                        if (!file.canRead()) {
                            throw new RuntimeException("Cannot read generated file: " + filePath);
                        }

                        // Чтение файла
                        byte[] content = Files.readAllBytes(file.toPath());

                        if (content == null || content.length == 0) {
                            throw new RuntimeException("Generated file is empty: " + filePath);
                        }

                        logger.info("Document generated successfully: {} ({} bytes)", file.getName(), content.length);

                        // УДАЛЯЕМ файл с сервера после чтения, чтобы не занимать место
                        boolean deleted = file.delete();
                        if (deleted) {
                            logger.debug("Temporary file deleted: {}", filePath);
                        } else {
                            logger.warn("Could not delete temporary file: {}", filePath);
                        }

                        // Формирование ответа с правильными заголовками для скачивания
                        String filename = file.getName();
                        return Mono.just(ResponseEntity.ok()
                                .header(HttpHeaders.CONTENT_DISPOSITION,
                                        "attachment; filename=\"" + filename + "\"")
                                .header(HttpHeaders.CONTENT_TYPE,
                                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                                .header("X-Filename", filename)
                                .contentLength(content.length)
                                .body(content));

                    } catch (Exception e) {
                        logger.error("Failed to generate document from URL: {}", cleanedUrl, e);
                        return Mono.error(new RuntimeException("Document generation failed: " + e.getMessage(), e));
                    }
                })
                .onErrorResume(e -> {
                    logger.error("Error in documentation generation from URL: {}", cleanedUrl, e);

                    String errorMessage;
                    HttpStatus status;

                    // Классификация ошибок
                    if (e instanceof java.util.concurrent.TimeoutException) {
                        errorMessage = "Request timeout: The API server took too long to respond";
                        status = HttpStatus.GATEWAY_TIMEOUT;
                    } else if (e.getCause() instanceof org.springframework.web.reactive.function.client.WebClientResponseException) {
                        var webClientEx = (org.springframework.web.reactive.function.client.WebClientResponseException) e.getCause();
                        if (webClientEx.getStatusCode() == HttpStatus.INTERNAL_SERVER_ERROR) {
                            errorMessage = "API server returned 500 Internal Server Error. The target API is experiencing issues.";
                            status = HttpStatus.BAD_GATEWAY;
                        } else {
                            errorMessage = String.format("API server error: HTTP %s - %s",
                                    webClientEx.getStatusCode(),
                                    webClientEx.getStatusText());
                            status = HttpStatus.BAD_GATEWAY;
                        }
                    } else if (e.getCause() instanceof java.net.UnknownHostException) {
                        errorMessage = "Cannot resolve host: " + cleanedUrl;
                        status = HttpStatus.BAD_REQUEST;
                    } else if (e.getCause() instanceof java.net.ConnectException) {
                        errorMessage = "Cannot connect to: " + cleanedUrl + ". Server may be down or port is closed.";
                        status = HttpStatus.BAD_GATEWAY;
                    } else {
                        errorMessage = "Error: " + e.getMessage();
                        status = HttpStatus.INTERNAL_SERVER_ERROR;
                    }

                    return Mono.just(ResponseEntity.status(status)
                            .contentType(MediaType.TEXT_PLAIN)
                            .body(errorMessage.getBytes()));
                })
                .timeout(Duration.ofSeconds(120)) // Увеличиваем таймаут до 2 минут
                .onErrorResume(java.util.concurrent.TimeoutException.class, e ->
                        Mono.just(ResponseEntity.status(HttpStatus.REQUEST_TIMEOUT)
                                .contentType(MediaType.TEXT_PLAIN)
                                .body("Error: Request timeout - operation took too long".getBytes()))
                );
    }

    // Дополнительный эндпоинт для проверки списка сгенерированных файлов
    @GetMapping("/list-files")
    public ResponseEntity<String> listGeneratedFiles() {
        try {
            File outputDir = new File("./generated-docs");
            if (!outputDir.exists()) {
                return ResponseEntity.ok("No generated files directory exists");
            }

            File[] files = outputDir.listFiles((dir, name) -> name.endsWith(".docx"));
            if (files == null || files.length == 0) {
                return ResponseEntity.ok("No DOCX files found in generated-docs directory");
            }

            StringBuilder sb = new StringBuilder("Generated files:\n");
            for (File file : files) {
                sb.append("- ").append(file.getName())
                        .append(" (").append(file.length()).append(" bytes)")
                        .append("\n");
            }

            return ResponseEntity.ok(sb.toString());
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Error listing files: " + e.getMessage());
        }
    }
}