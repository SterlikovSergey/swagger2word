package com.usnbook.swagger2word.service;

import com.usnbook.swagger2word.model.OpenApiSpec;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.reactive.function.client.WebClient;
import org.springframework.web.reactive.function.client.ExchangeStrategies;
import reactor.core.publisher.Mono;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.time.Duration;

@Service
public class ApiDocsService {

    private static final Logger logger = LoggerFactory.getLogger(ApiDocsService.class);

    private final WebClient webClient;

    public ApiDocsService(WebClient.Builder webClientBuilder) {
        // Увеличиваем лимит буфера до 10MB для обработки больших OpenAPI JSON
        ExchangeStrategies strategies = ExchangeStrategies.builder()
                .codecs(configurer -> configurer.defaultCodecs().maxInMemorySize(10 * 1024 * 1024))
                .build();

        this.webClient = webClientBuilder
                .exchangeStrategies(strategies)
                .build();
    }

    public Mono<OpenApiSpec> fetchApiDocs(String apiUrl) {
        logger.info("Fetching API docs from: {}", apiUrl);

        return webClient
                .get()
                .uri(apiUrl)
                .accept(MediaType.APPLICATION_JSON)
                .header("User-Agent", "Swagger2Word/1.0")
                .retrieve()
                .onStatus(status -> status.isError(), response -> {
                    // Получаем тело ошибки для лучшей диагностики
                    return response.bodyToMono(String.class)
                            .defaultIfEmpty("No error body provided by server")
                            .flatMap(errorBody -> {
                                String errorMsg = String.format("HTTP %s from %s. Error: %s",
                                        response.statusCode(), apiUrl, errorBody);
                                logger.error("API Server Error: {}", errorMsg);
                                return Mono.error(new RuntimeException(errorMsg));
                            });
                })
                .bodyToMono(OpenApiSpec.class)
                .timeout(Duration.ofSeconds(30))
                .doOnSuccess(spec -> {
                    if (spec != null) {
                        // Проверяем, не является ли ответ ошибкой в формате JSON
                        if (spec.getErrorMessage() != null) {
                            logger.warn("API returned error message: {}", spec.getErrorMessage());
                        }
                        logApiDocsInfo(spec);
                        logger.info("Successfully fetched API docs from: {}", apiUrl);
                    } else {
                        logger.warn("Received null API specification from: {}", apiUrl);
                    }
                })
                .doOnError(e -> {
                    logger.error("Failed to fetch API docs from: {}", apiUrl, e);

                    // Дополнительная диагностика для различных типов ошибок
                    if (e instanceof org.springframework.web.reactive.function.client.WebClientResponseException) {
                        var responseException = (org.springframework.web.reactive.function.client.WebClientResponseException) e;
                        logger.error("HTTP Status: {}, Headers: {}",
                                responseException.getStatusCode(),
                                responseException.getHeaders());
                    } else if (e instanceof java.net.UnknownHostException) {
                        logger.error("Unknown host: {}", apiUrl);
                    } else if (e instanceof java.net.ConnectException) {
                        logger.error("Connection refused: {}", apiUrl);
                    }
                })
                .onErrorResume(e -> {
                    String errorMessage;

                    if (e instanceof org.springframework.web.reactive.function.client.WebClientResponseException) {
                        var webClientEx = (org.springframework.web.reactive.function.client.WebClientResponseException) e;
                        errorMessage = String.format("API server returned error: HTTP %s - %s",
                                webClientEx.getStatusCode(),
                                webClientEx.getStatusText());
                    } else if (e.getCause() != null && e.getCause() instanceof java.net.UnknownHostException) {
                        errorMessage = String.format("Cannot resolve host: %s. Check the URL and network connectivity.", apiUrl);
                    } else if (e.getCause() != null && e.getCause() instanceof java.net.ConnectException) {
                        errorMessage = String.format("Cannot connect to: %s. Server may be down or port is closed.", apiUrl);
                    } else if (e instanceof java.util.concurrent.TimeoutException) {
                        errorMessage = String.format("Request timeout: Server %s took too long to respond.", apiUrl);
                    } else {
                        errorMessage = String.format("Failed to fetch API docs from: %s. Reason: %s",
                                apiUrl, e.getMessage());
                    }

                    logger.error(errorMessage);
                    return Mono.error(new RuntimeException(errorMessage, e));
                });
    }

    private void logApiDocsInfo(OpenApiSpec spec) {
        try {
            int pathsCount = 0;
            if (spec.getPaths() != null) {
                if (spec.getPaths() instanceof java.util.Map) {
                    pathsCount = ((java.util.Map<?, ?>) spec.getPaths()).size();
                }
            }

            logger.info("API Documentation Info - Title: {}, Version: {}, Paths: {}",
                    spec.getInfo() != null ? spec.getInfo().getTitle() : "N/A",
                    spec.getInfo() != null ? spec.getInfo().getVersion() : "N/A",
                    pathsCount);

        } catch (Exception e) {
            logger.warn("Error while logging API docs info: {}", e.getMessage());
        }
    }
}