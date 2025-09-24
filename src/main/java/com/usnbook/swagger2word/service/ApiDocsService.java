package com.usnbook.swagger2word.service;

import com.usnbook.swagger2word.model.OpenApiSpec;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.core.publisher.Mono;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;

@Service
public class ApiDocsService {

    private static final Logger logger = LoggerFactory.getLogger(ApiDocsService.class);

    private final WebClient webClient;
    private final String apiDocsUrl;

    public ApiDocsService(WebClient.Builder webClientBuilder,
                          @Value("${app.api-docs-url}") String apiDocsUrl) {
        this.webClient = webClientBuilder
                .baseUrl(apiDocsUrl)
                .defaultHeader(HttpHeaders.ACCEPT, MediaType.APPLICATION_JSON_VALUE)
                .build();
        this.apiDocsUrl = apiDocsUrl;
    }

    public Mono<OpenApiSpec> fetchApiDocs() {
        logger.info("Fetching API docs from: {}", apiDocsUrl);

        return webClient.get()
                .uri(apiDocsUrl)
                .retrieve()
                .bodyToMono(OpenApiSpec.class)
                .doOnSuccess(this::logApiDocsInfo)
                .doOnError(e -> logger.error("Failed to fetch API docs from: {}", apiDocsUrl, e))
                .onErrorMap(e -> new RuntimeException("Failed to fetch API docs from: " + apiDocsUrl, e));
    }

    private void logApiDocsInfo(OpenApiSpec spec) {
        int pathsCount = 0;
        if (spec.getPaths() != null) {
            // Если paths - это Map
            if (spec.getPaths() instanceof Map) {
                pathsCount = ((Map<?, ?>) spec.getPaths()).size();
            }
            // Если paths - это объект Paths (новая модель)
            else if (spec.getPaths() != null) {
                try {
                    java.lang.reflect.Method getPathsMethod = spec.getPaths().getClass().getMethod("getPaths");
                    Object pathsMap = getPathsMethod.invoke(spec.getPaths());
                    if (pathsMap instanceof Map) {
                        pathsCount = ((Map<?, ?>) pathsMap).size();
                    }
                } catch (Exception e) {
                    logger.debug("Could not extract paths count via reflection: {}", e.getMessage());
                }
            }
        }

        logger.info("Successfully fetched API docs, {} paths found", pathsCount);

        // Дополнительная диагностика
        if (spec.getPaths() != null) {
            logger.debug("Paths type: {}", spec.getPaths().getClass().getSimpleName());
            if (spec.getPaths() instanceof Map) {
                logger.debug("Paths is Map with {} entries", ((Map<?, ?>) spec.getPaths()).size());
            }
        }
    }
}