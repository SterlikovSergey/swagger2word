# Swagger2Word

[![Java](https://img.shields.io/badge/Java-11%2B-blue)](https://www.oracle.com/java/) [![Spring Boot](https://img.shields.io/badge/Spring%20Boot-3.0%2B-green)](https://spring.io/projects/spring-boot) [![Gradle](https://img.shields.io/badge/Gradle-8.0%2B-orange)](https://gradle.org/) [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Swagger2Word** — микросервис для автоматической генерации профессиональной документации API в формате Microsoft Word (.docx) на основе OpenAPI спецификаций (Swagger). Сервис универсален: поддерживает любые API-docs endpoints (например, `/v3/api-docs` или `/api-docs`), нормализует URL и генерирует структурированный документ с разделами (титульная страница, общая информация, серверы, группы API, endpoints, модели данных).

Сервис следует принципам **S.O.L.I.D.** для модульности и расширяемости. Поддерживает несколько источников API через параметр URL.

## Описание

Swagger2Word извлекает OpenAPI спецификацию из удаленного источника, парсит её и генерирует Word-документ с:
- Титульной страницей (название, описание, версия).
- Общей информацией (серверы, теги/группы).
- Детальным описанием endpoints (методы, параметры, тела запросов, ответы).
- Моделями данных (схемы с свойствами, типами, обязательностью).

**Ключевые особенности:**
- **Универсальность**: Работает с любым OpenAPI 3.0+ endpoint (автоматическая нормализация URL).
- **Конфигурация**: Таймауты, разрешенные домены, директория вывода через `application.yml`.
- **Расширяемость**: Модульные форматтеры для добавления новых разделов (Open-Closed Principle).
- **Безопасность**: Валидация URL и доменов.
- **Reactive**: Использует WebFlux для асинхронного fetch.

## Установка

### Требования
- Java 11+
- Gradle 8.0+
- Spring Boot 3.0+

### Клонирование и сборка
```bash
git clone <your-repo-url>
cd swagger2word
gradle build
```

### Настройка Sensitive Данных
Создайте файл `src/main/resources/secret.yml` (он будет автоматически добавлен в `.gitignore`):
```yaml
app:
  swagger2word:
    default-api-docs-url: https://your-sensitive-api.example.com/api-docs  # Default URL
    allowed-domains:
      - "your-sensitive-domain1.example.com"
      - "your-sensitive-domain2.example.com"
```
- `secret.yml` переопределяет `application.yml` для sensitive данных (URL, домены).
- Пример: Используйте для хранения приватных API endpoints, которые не должны попадать в репозиторий.

### Запуск
```bash
gradle bootRun
```
Сервис запустится на `http://localhost:8081`.

## Использование

### Генерация документа
**Default (из конфигурации)**:
```
GET http://localhost:8081/api/generate-doc
```
- Использует `default-api-docs-url` из `secret.yml` или `application.yml`.

**Универсальный (произвольный URL)**:
```
GET http://localhost:8081/api/generate-doc?url=https://your-api-host.example.com/api-docs
```
- `url` - полный URL до OpenAPI JSON (нормализуется автоматически, e.g., добавит `/api-docs` если нужно).

**С параметрами**:
```
GET http://localhost:8081/api/generate-doc?url=https://your-api-host.example.com/v3/api-docs&includeDiagnostics=true
```
- `includeDiagnostics`: Включить диагностику в документ (default: false).

**Ответ**:
- HTTP 200 с .docx в body (Content-Disposition: attachment).
- Ошибки: HTTP 400/500 с JSON-сообщением.

### Другие Endpoints
- `GET /api/generate-doc/supported-domains`: Список разрешенных доменов (из конфигурации).
- `GET /api/generate-doc/health`: Health-check.

## Конфигурация

### application.yml (основная)
```yaml
server:
  port: 8081
  max-http-request-size: 10MB  # Для больших OpenAPI JSON
  max-http-request-header-size: 65536

app:
  swagger2word:
    # Default URL (переопределяется secret.yml)
    default-api-docs-url: https://default-api.example.com/api-docs
    default-output-directory: ./generated-docs
    allowed-domains: []  # [] = все; укажите в secret.yml для безопасности
    connection-timeout-ms: 10000
    read-timeout-ms: 30000
    enable-diagnostics: true

spring:
  web:
    resources:
      add-mappings: false
  webflux:
    max-in-memory-size: 10MB
    codec:
      max-in-memory-size: 10MB
  jackson:
    default-property-inclusion: non_null
    deserialization:
      fail-on-unknown-properties: false  # Игнор лишних полей в OpenAPI

logging:
  level:
    com.usnbook.swagger2word: INFO
    org.springframework.web.reactive.function.client: DEBUG  # Для отладки WebClient
```

### secret.yml (sensitive данные, в .gitignore)
```yaml
app:
  swagger2word:
    default-api-docs-url: https://your-private-api.example.com/api-docs  # Приватный default
    allowed-domains:
      - "private-domain1.example.com"
      - "private-domain2.example.com"
```
- **Важно**: `secret.yml` загружается Spring Boot автоматически (если в `src/main/resources`), но не коммитится в Git.

- **allowed-domains**: Пустой = все URL разрешены. Укажите для безопасности (только в secret.yml).
- **default-api-docs-url**: Для default-запроса без параметра.

## Структура Проекта

```
src/main/java/com/usnbook/swagger2word/
├── Swagger2wordApplication.java      # Главный класс
├── config/
│   └── Swagger2WordProperties.java   # Конфигурация
├── controller/
│   └── DocumentationController.java  # Универсальный контроллер
├── model/
│   └── OpenApiSpec.java              # Модель OpenAPI
└── service/
    ├── ApiDocsService.java           # Fetch API-docs (универсальный)
    ├── WordDocumentService.java       # Генерация Word
    ├── validation/
    │   └── UrlValidator.java          # Валидация URL
    └── api/
        └── ApiSpecificationProvider.java  # Интерфейс для fetch (DIP)
```

## S.O.L.I.D. Принципы в Реализации

- **Single Responsibility (SRP)**: `ApiDocsService` - только fetch, `WordDocumentService` - только генерация, `UrlValidator` - только валидация.
- **Open-Closed (OCP)**: Добавьте новые валидаторы или сервисы через `@Service` без изменения кода.
- **Liskov Substitution (LSP)**: `DefaultUrlValidator` заменяем на любой `UrlValidator`.
- **Interface Segregation (ISP)**: `ApiSpecificationProvider` - только fetch, без лишних методов.
- **Dependency Inversion (DIP)**: Контроллер зависит от интерфейсов (`ApiSpecificationProvider`), не от реализаций.

## Примеры

### Пример 1: Default (из конфигурации)
```
curl -X GET "http://localhost:8081/api/generate-doc" \
  -H "Accept: application/octet-stream" \
  --output default-doc.docx
```
- Скачает .docx на основе `default-api-docs-url`.

### Пример 2: Произвольный API
```
curl -X GET "http://localhost:8081/api/generate-doc?url=https://api.example.com/v3/api-docs" \
  -H "Accept: application/octet-stream" \
  --output custom-doc.docx
```
- Нормализует URL, генерирует `custom-doc.docx`.

### Пример 3: С диагностикой
```
curl -X GET "http://localhost:8081/api/generate-doc?url=https://api.example.com/api-docs&includeDiagnostics=true" \
  -H "Accept: application/octet-stream" \
  --output doc-with-diag.docx
```

## Troubleshooting

- **"Invalid URL"**: Проверьте `allowed-domains` в secret.yml или протокол (http/https).
- **"Timeout"**: Увеличьте `read-timeout-ms` в yml.
- **"No paths found"**: Убедитесь, что API-docs доступен (проверьте curl на JSON).
- **"File not found"**: Проверьте права на `default-output-directory`.
- **Логи**: Установите `logging.level.com.usnbook.swagger2word=DEBUG` для детальной отладки.

## Контакты
- Автор: Стерликов Сергей
- Репозиторий: https://github.com/SterlikovSergey

---

*Generated on September 24, 2025*