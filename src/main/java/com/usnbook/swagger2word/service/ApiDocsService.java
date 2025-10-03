package com.usnbook.swagger2word.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.core.publisher.Mono;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Service
public class ApiDocsService {

    private static final Logger logger = LoggerFactory.getLogger(ApiDocsService.class);

    private final WebClient.Builder webClientBuilder;
    private final ObjectMapper objectMapper;

    public ApiDocsService(WebClient.Builder webClientBuilder, ObjectMapper objectMapper) {
        this.webClientBuilder = webClientBuilder;
        this.objectMapper = objectMapper;
    }

    public Mono<JsonNode> fetchApiDocs(String apiUrl) {
        logger.info("Fetching API docs from: {}", apiUrl);

        return webClientBuilder.build()
                .get()
                .uri(apiUrl)
                .accept(MediaType.APPLICATION_JSON)
                .retrieve()
                .bodyToMono(String.class)
                .map(jsonString -> {
                    try {
                        JsonNode rootNode = objectMapper.readTree(jsonString);
                        validateOpenApiSpec(rootNode);
                        return rootNode;
                    } catch (Exception e) {
                        logger.error("Failed to parse JSON from: {}", apiUrl, e);
                        throw new RuntimeException("Failed to parse JSON response: " + e.getMessage(), e);
                    }
                })
                .doOnSuccess(spec -> logApiDocsInfo(spec))
                .doOnError(e -> logger.error("Failed to fetch API docs from: {}", apiUrl, e));
    }

    private void validateOpenApiSpec(JsonNode spec) {
        if (!spec.has("openapi")) {
            throw new RuntimeException("Invalid OpenAPI specification: missing 'openapi' field");
        }

        if (!spec.has("info")) {
            throw new RuntimeException("Invalid OpenAPI specification: missing 'info' field");
        }

        String openapiVersion = spec.path("openapi").asText();
        if (!openapiVersion.startsWith("3.")) {
            logger.warn("OpenAPI version {} may not be fully supported", openapiVersion);
        }
    }

    private void logApiDocsInfo(JsonNode spec) {
        try {
            JsonNode infoNode = spec.path("info");
            String title = infoNode.path("title").asText("Unknown");
            String version = infoNode.path("version").asText("Unknown");

            int pathsCount = 0;
            JsonNode pathsNode = spec.path("paths");
            if (pathsNode.isObject()) {
                pathsCount = pathsNode.size();
            }

            int schemasCount = 0;
            JsonNode componentsNode = spec.path("components");
            if (componentsNode.isObject()) {
                JsonNode schemasNode = componentsNode.path("schemas");
                if (schemasNode.isObject()) {
                    schemasCount = schemasNode.size();
                }
            }

            logger.info("Successfully fetched API docs: '{}' (v{}), {} paths, {} schemas",
                    title, version, pathsCount, schemasCount);

        } catch (Exception e) {
            logger.warn("Error logging API docs info: {}", e.getMessage());
        }
    }
}