package com.usnbook.swagger2word.model;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;

import java.util.*; // ИСПРАВЛЕНИЕ: Добавлен импорт коллекций

@Data
@JsonIgnoreProperties(ignoreUnknown = true)
public class OpenApiSpec {
    private String openapi;
    private Info info;
    private List<Server> servers;
    private List<Tag> tags;
    private Map<String, Path> paths; // ИСПРАВЛЕНИЕ: Оставляем Map для обратной совместимости
    private Components components;
    private ExternalDocs externalDocs;

    @Data
    public static class Info {
        private String title;
        private String description;
        private String version;
        private Contact contact;
        private License license;
    }

    @Data
    public static class Contact {
        private String name;
        private String url;
        private String email;
    }

    @Data
    public static class License {
        private String name;
        private String url;
    }

    @Data
    public static class Server {
        private String url;
        private String description;
    }

    @Data
    public static class Tag {
        private String name;
        private String description;
        private ExternalDocs externalDocs;
    }

    @Data
    public static class ExternalDocs {
        private String description;
        private String url;
    }

    // ИСПРАВЛЕНИЕ: Расширенный класс Path с поддержкой всех HTTP методов
    @Data
    public static class Path {
        @JsonProperty("get")
        private Operation getOperation;

        @JsonProperty("put")
        private Operation putOperation;

        @JsonProperty("post")
        private Operation postOperation;

        @JsonProperty("delete")
        private Operation deleteOperation;

        @JsonProperty("patch")
        private Operation patchOperation;

        @JsonProperty("head")
        private Operation headOperation;

        @JsonProperty("options")
        private Operation optionsOperation;

        @JsonProperty("trace")
        private Operation traceOperation;

        // Для обратной совместимости - если используется Map
        private Map<String, Operation> operations;

        // ИСПРАВЛЕНИЕ: Универсальный метод получения всех операций
        public Map<String, Operation> getAllOperations() {
            Map<String, Operation> allOps = new LinkedHashMap<>();

            // Вариант 1: Через operations Map (если используется)
            if (this.operations != null && !this.operations.isEmpty()) {
                allOps.putAll(this.operations);
            }

            // Вариант 2: Через отдельные поля (если Map пустой или отсутствует)
            if (allOps.isEmpty()) {
                if (getOperation != null) allOps.put("get", getOperation);
                if (putOperation != null) allOps.put("put", putOperation);
                if (postOperation != null) allOps.put("post", postOperation);
                if (deleteOperation != null) allOps.put("delete", deleteOperation);
                if (patchOperation != null) allOps.put("patch", patchOperation);
                if (headOperation != null) allOps.put("head", headOperation);
                if (optionsOperation != null) allOps.put("options", optionsOperation);
                if (traceOperation != null) allOps.put("trace", traceOperation);
            }

            return allOps.isEmpty() ? null : allOps;
        }

        // Устаревший метод для совместимости
        public Map<String, Operation> getOperations() {
            Map<String, Operation> ops = getAllOperations();
            return ops != null ? ops : new HashMap<>();
        }
    }

    @Data
    public static class Operation {
        private List<String> tags;
        private String summary;
        private String description;
        private String operationId;
        private List<Parameter> parameters;
        private RequestBody requestBody;
        private Map<String, Response> responses;
        private Boolean deprecated;
        private List<SecurityRequirement> security;
        private List<Map<String, List<String>>> securityRequirements;
    }

    @Data
    public static class SecurityRequirement {
        private String name;
        private List<String> scopes;
    }

    @Data
    public static class Parameter {
        private String name;
        private String in;
        private String description;
        private Boolean required;
        private Boolean deprecated;
        private Boolean allowEmptyValue;
        private Schema schema;
        private Object example;
        private Map<String, Example> examples;
        private Content content;
    }

    @Data
    public static class RequestBody {
        private String description;
        private Boolean required;
        private Map<String, MediaType> content;
    }

    @Data
    public static class MediaType {
        private Schema schema;
        private Object example;
        private Map<String, Example> examples;
        private Map<String, Encoding> encoding;
    }

    @Data
    public static class Encoding {
        private String contentType;
        private Map<String, Header> headers;
        private Boolean explode;
        private Boolean allowReserved;
    }

    @Data
    public static class Example {
        private String summary;
        private String description;
        private Object value;
        private String externalValue;
    }

    @Data
    public static class Response {
        private String description;
        private Map<String, MediaType> content;
        private Map<String, Header> headers;
        private Map<String, Link> links;
    }

    @Data
    public static class Header {
        private String description;
        private Boolean required;
        private Boolean deprecated;
        private Schema schema;
        private Object example;
    }

    @Data
    public static class Link {
        private String operationRef;
        private String operationId;
        private Map<String, Object> parameters;
        private Object requestBody;
        private String description;
    }

    @Data
    public static class Schema {
        private String type;
        // Геттер для совместимости
        private String ref;
        private String format;
        private String title;
        private String description;
        private Object example;
        private Map<String, Schema> properties;
        private Schema items;
        private List<String> required;
        private Boolean nullable;
        private Boolean readOnly;
        private Boolean writeOnly;
        private Boolean deprecated;
        private Integer minLength;
        private Integer maxLength;
        private String pattern;
        private Integer minimum;
        private Integer maximum;
        private Integer minItems;
        private Integer maxItems;
        private Boolean uniqueItems;
        private String discriminator;
        private Map<String, String> enumValues;
        private Object defaultValue;
        private Schema additionalProperties;
        private List<Schema> allOf;
        private List<Schema> oneOf;
        private List<Schema> anyOf;
        private Schema not;

    }

    @Data
    public static class Components {
        private Map<String, Schema> schemas;
        private Map<String, Response> responses;
        private Map<String, Parameter> parameters;
        private Map<String, Example> examples;
        private Map<String, RequestBody> requestBodies;
        private Map<String, Header> headers;
        private Map<String, SecurityScheme> securitySchemes;
        private Map<String, Link> links;
    }

    @Data
    public static class SecurityScheme {
        private String type;
        private String description;
        private String name;
        private String in;
        private String scheme;
        private String bearerFormat;
        private String openIdConnectUrl;
        private Map<String, String> flows;
    }

    @Data
    public static class Content {
        private Map<String, MediaType> mediaTypes;
    }
}