package com.usnbook.swagger2word.service;

import com.fasterxml.jackson.databind.JsonNode;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.math.BigInteger;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;

@Service
public class WordDocumentService {

    private static final Logger logger = LoggerFactory.getLogger(WordDocumentService.class);

    public byte[] generateWordDocumentFromJson(JsonNode apiSpec) throws Exception {
        if (apiSpec == null) {
            throw new IllegalArgumentException("API specification cannot be null");
        }

        try (XWPFDocument document = new XWPFDocument()) {
            // Настройка документа
            setupDocumentProperties(document);

            // Генерация содержимого
            addTitlePage(document, apiSpec);
            addTableOfContents(document, apiSpec);
            addGeneralInfo(document, apiSpec);
            addServersSection(document, apiSpec);
            addTagsSection(document, apiSpec);
            addEndpointsSection(document, apiSpec);
            addSchemasSection(document, apiSpec);
            addSecuritySchemesSection(document, apiSpec);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            document.write(baos);
            return baos.toByteArray();
        }
    }

    private void setupDocumentProperties(XWPFDocument document) {
        // Установка свойств документа
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageSz pageSz = sectPr.addNewPgSz();
        pageSz.setW(BigInteger.valueOf(11906)); // A4 width in twips
        pageSz.setH(BigInteger.valueOf(16838)); // A4 height in twips

        CTPageMar pageMar = sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(1800));  // 1.5 cm
        pageMar.setRight(BigInteger.valueOf(1800)); // 1.5 cm
        pageMar.setTop(BigInteger.valueOf(1440));   // 2.54 cm
        pageMar.setBottom(BigInteger.valueOf(1440)); // 2.54 cm
        pageMar.setHeader(BigInteger.valueOf(720)); // 1.27 cm
        pageMar.setFooter(BigInteger.valueOf(720)); // 1.27 cm
    }

    private void addTitlePage(XWPFDocument document, JsonNode apiSpec) {
        JsonNode infoNode = apiSpec.path("info");

        // Главный заголовок
        XWPFParagraph titleParagraph = document.createParagraph();
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        titleParagraph.setSpacingBefore(1200);
        titleParagraph.setSpacingAfter(400);

        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText("ДОКУМЕНТАЦИЯ API");
        titleRun.setBold(true);
        titleRun.setFontSize(28);
        titleRun.setFontFamily("Arial");
        titleRun.setColor("1F4E79");
        titleRun.addBreak();

        // Название API
        String apiTitle = infoNode.path("title").asText("Unnamed API");
        XWPFRun subtitleRun = titleParagraph.createRun();
        subtitleRun.setText(apiTitle);
        subtitleRun.setBold(true);
        subtitleRun.setFontSize(20);
        subtitleRun.setFontFamily("Arial");
        subtitleRun.setColor("2E75B6");
        subtitleRun.addBreak();

        // Описание
        if (infoNode.has("description")) {
            String description = infoNode.path("description").asText("").trim();
            if (!description.isEmpty()) {
                XWPFParagraph descParagraph = document.createParagraph();
                descParagraph.setAlignment(ParagraphAlignment.CENTER);
                descParagraph.setSpacingAfter(300);
                descParagraph.setSpacingBefore(200);
                descParagraph.setIndentationLeft(720);
                descParagraph.setIndentationRight(720);

                XWPFRun descRun = descParagraph.createRun();
                descRun.setText(description);
                descRun.setFontSize(12);
                descRun.setFontFamily("Arial");
                descRun.setColor("333333");
            }
        }

        // Контактная информация
        if (infoNode.has("contact")) {
            addContactInfo(document, infoNode.path("contact"));
        }

        // Лицензия
        if (infoNode.has("license")) {
            addLicenseInfo(document, infoNode.path("license"));
        }

        // Метаданные
        addMetadata(document, apiSpec);
        addPageBreak(document);
    }

    private void addContactInfo(XWPFDocument document, JsonNode contactNode) {
        if (contactNode.isMissingNode()) return;

        List<String> contactLines = new ArrayList<>();
        if (contactNode.has("name")) {
            contactLines.add("Контакты: " + contactNode.path("name").asText());
        }
        if (contactNode.has("email")) {
            contactLines.add("Email: " + contactNode.path("email").asText());
        }
        if (contactNode.has("url")) {
            contactLines.add("URL: " + contactNode.path("url").asText());
        }

        if (!contactLines.isEmpty()) {
            XWPFParagraph contactParagraph = document.createParagraph();
            contactParagraph.setAlignment(ParagraphAlignment.CENTER);
            contactParagraph.setSpacingBefore(150);
            contactParagraph.setSpacingAfter(150);

            XWPFRun contactRun = contactParagraph.createRun();
            contactRun.setText(String.join(" | ", contactLines));
            contactRun.setFontSize(10);
            contactRun.setFontFamily("Arial");
            contactRun.setColor("555555");
        }
    }

    private void addLicenseInfo(XWPFDocument document, JsonNode licenseNode) {
        if (licenseNode.isMissingNode()) return;

        XWPFParagraph licenseParagraph = document.createParagraph();
        licenseParagraph.setAlignment(ParagraphAlignment.CENTER);
        licenseParagraph.setSpacingAfter(150);

        XWPFRun licenseRun = licenseParagraph.createRun();
        licenseRun.setText("Лицензия: " + licenseNode.path("name").asText("Unknown"));

        if (licenseNode.has("url")) {
            licenseRun.addBreak();
            licenseRun.setText("URL: " + licenseNode.path("url").asText());
        }

        licenseRun.setFontSize(10);
        licenseRun.setFontFamily("Arial");
        licenseRun.setColor("555555");
    }

    private void addMetadata(XWPFDocument document, JsonNode apiSpec) {
        JsonNode infoNode = apiSpec.path("info");

        XWPFParagraph metaParagraph = document.createParagraph();
        metaParagraph.setAlignment(ParagraphAlignment.CENTER);
        metaParagraph.setSpacingAfter(600);
        metaParagraph.setSpacingBefore(300);

        XWPFRun metaRun = metaParagraph.createRun();
        metaRun.setText("Версия API: " + infoNode.path("version").asText("N/A"));
        metaRun.setFontSize(10);
        metaRun.setFontFamily("Arial");
        metaRun.setColor("666666");
        metaRun.addBreak();

        metaRun.setText("Спецификация OpenAPI: " + apiSpec.path("openapi").asText("N/A"));
        metaRun.addBreak();

        metaRun.setText("Сгенерировано: " +
                LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm:ss")));
    }

    private void addTableOfContents(XWPFDocument document, JsonNode apiSpec) {
        XWPFParagraph tocTitle = document.createParagraph();
        tocTitle.setAlignment(ParagraphAlignment.CENTER);
        tocTitle.setSpacingAfter(300);

        XWPFRun tocTitleRun = tocTitle.createRun();
        tocTitleRun.setText("СОДЕРЖАНИЕ");
        tocTitleRun.setBold(true);
        tocTitleRun.setFontSize(16);
        tocTitleRun.setFontFamily("Arial");
        tocTitleRun.setColor("1F4E79");

        // Основные разделы
        addTocItem(document, "1. ОБЩАЯ ИНФОРМАЦИЯ");

        if (hasServers(apiSpec)) {
            addTocItem(document, "2. СЕРВЕРЫ");
        }

        if (hasTags(apiSpec)) {
            addTocItem(document, "3. ГРУППЫ API");
        }

        addTocItem(document, "4. ENDPOINTS");

        // Подразделы endpoints
        Map<String, List<EndpointInfo>> groupedEndpoints = groupEndpointsByTags(apiSpec);
        int groupNum = 1;
        for (String groupName : groupedEndpoints.keySet()) {
            addTocItem(document, "   4." + groupNum + ". " + groupName);
            groupNum++;
        }

        if (hasSchemas(apiSpec)) {
            addTocItem(document, "5. МОДЕЛИ ДАННЫХ");
        }

        if (hasSecuritySchemes(apiSpec)) {
            addTocItem(document, "6. БЕЗОПАСНОСТЬ");
        }

        addPageBreak(document);
    }

    private void addTocItem(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(80);

        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily("Arial");
        run.setFontSize(11);
        run.setColor("333333");
    }

    private void addGeneralInfo(XWPFDocument document, JsonNode apiSpec) {
        JsonNode infoNode = apiSpec.path("info");

        addSectionTitle(document, "1. ОБЩАЯ ИНФОРМАЦИЯ", 18, "1F4E79");

        XWPFTable infoTable = document.createTable(5, 2);
        infoTable.setWidth("100%");

        // Настройка ширины колонок
        setTableColumnWidths(infoTable, new int[]{35, 65});

        addTableRow(infoTable, "Название API", infoNode.path("title").asText("N/A"), 0);
        addTableRow(infoTable, "Версия API", infoNode.path("version").asText("N/A"), 1);
        addTableRow(infoTable, "Спецификация OpenAPI", apiSpec.path("openapi").asText("N/A"), 2);

        String description = infoNode.path("description").asText("").trim();
        if (!description.isEmpty()) {
            addTableRow(infoTable, "Описание", description, 3);
        } else {
            infoTable.getRow(3).getCell(0).setText("");
            infoTable.getRow(3).getCell(1).setText("");
        }

        // Термины и условия
        if (infoNode.has("termsOfService")) {
            addTableRow(infoTable, "Условия использования", infoNode.path("termsOfService").asText(), 4);
        }

        addSectionSpacing(document, 400);
    }

    private void addServersSection(XWPFDocument document, JsonNode apiSpec) {
        if (!hasServers(apiSpec)) {
            return;
        }

        addSectionTitle(document, "2. СЕРВЕРЫ", 16, "1F4E79");

        JsonNode servers = apiSpec.get("servers");
        for (int i = 0; i < servers.size(); i++) {
            JsonNode server = servers.get(i);
            addServerInfo(document, i + 1, server);
        }

        addSectionSpacing(document, 400);
    }

    private void addServerInfo(XWPFDocument document, int serverNum, JsonNode server) {
        XWPFParagraph serverParagraph = document.createParagraph();
        serverParagraph.setAlignment(ParagraphAlignment.LEFT);
        serverParagraph.setSpacingAfter(120);

        XWPFRun serverRun = serverParagraph.createRun();
        serverRun.setFontFamily("Arial");
        serverRun.setFontSize(11);

        // Номер и URL сервера
        serverRun.setText(serverNum + ". ");
        serverRun.setBold(true);
        serverRun.setColor("2E75B6");
        serverRun.setText(server.path("url").asText("N/A"));
        serverRun.setBold(false);
        serverRun.setColor("333333");

        // Описание сервера
        if (server.has("description")) {
            String description = server.path("description").asText("").trim();
            if (!description.isEmpty()) {
                serverRun.addBreak();
                serverRun.setText("   Описание: " + description);
                serverRun.setItalic(true);
            }
        }

        // Переменные сервера
        if (server.has("variables")) {
            JsonNode variables = server.get("variables");
            if (variables.size() > 0) {
                serverRun.addBreak();
                serverRun.setText("   Переменные:");
                serverRun.setItalic(false);

                Iterator<String> varNames = variables.fieldNames();
                while (varNames.hasNext()) {
                    String varName = varNames.next();
                    JsonNode variable = variables.get(varName);
                    serverRun.addBreak();
                    serverRun.setText("     - " + varName + ": " +
                            variable.path("default").asText("") +
                            (variable.has("description") ?
                                    " (" + variable.path("description").asText() + ")" : ""));
                }
            }
        }
    }

    private void addTagsSection(XWPFDocument document, JsonNode apiSpec) {
        if (!hasTags(apiSpec)) {
            return;
        }

        addSectionTitle(document, "3. ГРУППЫ API", 16, "1F4E79");

        JsonNode tags = apiSpec.get("tags");
        for (int i = 0; i < tags.size(); i++) {
            JsonNode tag = tags.get(i);
            addTagInfo(document, i + 1, tag);
        }

        addSectionSpacing(document, 400);
    }

    private void addTagInfo(XWPFDocument document, int tagNum, JsonNode tag) {
        XWPFParagraph tagParagraph = document.createParagraph();
        tagParagraph.setAlignment(ParagraphAlignment.LEFT);
        tagParagraph.setSpacingAfter(100);

        XWPFRun tagRun = tagParagraph.createRun();
        tagRun.setFontFamily("Arial");
        tagRun.setFontSize(11);

        tagRun.setText(tagNum + ". ");
        tagRun.setBold(true);
        tagRun.setColor("2E75B6");
        tagRun.setText(tag.path("name").asText("Unnamed"));
        tagRun.setBold(false);
        tagRun.setColor("333333");

        if (tag.has("description")) {
            String description = tag.path("description").asText("").trim();
            if (!description.isEmpty()) {
                tagRun.addBreak();
                tagRun.setText("   " + description);
                tagRun.setItalic(true);
            }
        }

        // Внешняя документация
        if (tag.has("externalDocs")) {
            JsonNode externalDocs = tag.get("externalDocs");
            tagRun.addBreak();
            tagRun.setItalic(false);
            tagRun.setText("   Документация: ");
            tagRun.setColor("0066CC");
            if (externalDocs.has("url")) {
                tagRun.setText(externalDocs.path("url").asText());
            }
        }
    }

    private void addEndpointsSection(XWPFDocument document, JsonNode apiSpec) {
        addSectionTitle(document, "4. ENDPOINTS", 20, "1F4E79");

        Map<String, List<EndpointInfo>> groupedEndpoints = groupEndpointsByTags(apiSpec);

        if (groupedEndpoints.isEmpty()) {
            addEmptyEndpointsMessage(document);
        } else {
            displayGroupedEndpoints(document, groupedEndpoints);
        }

        addSectionSpacing(document, 400);
    }

    private Map<String, List<EndpointInfo>> groupEndpointsByTags(JsonNode apiSpec) {
        Map<String, List<EndpointInfo>> grouped = new LinkedHashMap<>();

        if (!apiSpec.has("paths")) {
            return grouped;
        }

        JsonNode paths = apiSpec.get("paths");
        Iterator<String> pathNames = paths.fieldNames();

        while (pathNames.hasNext()) {
            String path = pathNames.next();
            JsonNode pathItem = paths.get(path);

            // Обрабатываем все HTTP методы
            for (String method : Arrays.asList("get", "post", "put", "delete", "patch", "head", "options", "trace")) {
                if (pathItem.has(method)) {
                    JsonNode operation = pathItem.get(method);
                    EndpointInfo endpoint = new EndpointInfo(path, method.toUpperCase(), operation);

                    // Группируем по тегам
                    if (operation.has("tags") && operation.get("tags").size() > 0) {
                        for (JsonNode tag : operation.get("tags")) {
                            String tagName = tag.asText();
                            grouped.computeIfAbsent(tagName, k -> new ArrayList<>()).add(endpoint);
                        }
                    } else {
                        grouped.computeIfAbsent("Без категории", k -> new ArrayList<>()).add(endpoint);
                    }
                }
            }
        }

        return grouped;
    }

    private void displayGroupedEndpoints(XWPFDocument document, Map<String, List<EndpointInfo>> groupedEndpoints) {
        int groupNum = 1;
        for (Map.Entry<String, List<EndpointInfo>> entry : groupedEndpoints.entrySet()) {
            addEndpointGroup(document, groupNum, entry.getKey(), entry.getValue());
            groupNum++;
        }
    }

    private void addEndpointGroup(XWPFDocument document, int groupNum, String groupName, List<EndpointInfo> endpoints) {
        addSubsectionTitle(document, "4." + groupNum + ". " + groupName.toUpperCase(), 16, "2E75B6");

        // Сортируем endpoints
        endpoints.sort(Comparator.comparing(EndpointInfo::getPath)
                .thenComparing(EndpointInfo::getMethod));

        for (EndpointInfo endpoint : endpoints) {
            addDetailedEndpointTable(document, endpoint);
            addSectionSpacing(document, 100);
        }
    }

    private void addDetailedEndpointTable(XWPFDocument document, EndpointInfo endpoint) {
        JsonNode operation = endpoint.getOperation();

        // Основная таблица endpoint
        XWPFTable mainTable = document.createTable(1, 2);
        mainTable.setWidth("100%");
        setTableColumnWidths(mainTable, new int[]{20, 80});

        // Заголовок с методом и путем
        XWPFTableRow headerRow = mainTable.getRow(0);
        headerRow.getCell(0).setText("Метод");

        String methodColor = getMethodColor(endpoint.getMethod());
        XWPFTableCell methodCell = headerRow.getCell(1);
        methodCell.setText(endpoint.getMethod() + " " + endpoint.getPath());
        methodCell.setColor(methodColor);
        styleTableHeader(headerRow);

        if (operation.has("summary") && !operation.path("summary").asText("").trim().isEmpty()) {
            addTableRow(mainTable, "Краткое описание", operation.path("summary").asText(), mainTable.getNumberOfRows());
        }

        if (operation.has("description") && !operation.path("description").asText("").trim().isEmpty()) {
            addTableRow(mainTable, "Описание", operation.path("description").asText(), mainTable.getNumberOfRows());
        }

        if (operation.has("operationId") && !operation.path("operationId").asText("").trim().isEmpty()) {
            addTableRow(mainTable, "ID операции", operation.path("operationId").asText(), mainTable.getNumberOfRows());
        }

        // Параметры
        if (operation.has("parameters") && operation.get("parameters").size() > 0) {
            addDetailedParametersSection(mainTable, operation.get("parameters"));
        }

        // Тело запроса
        if (operation.has("requestBody")) {
            addDetailedRequestBodySection(mainTable, operation.get("requestBody"));
        }

        // Ответы
        if (operation.has("responses") && operation.get("responses").size() > 0) {
            addDetailedResponsesSection(mainTable, operation.get("responses"));
        }

        // Безопасность
        if (operation.has("security") && operation.get("security").size() > 0) {
            addSecurityRequirementsSection(mainTable, operation.get("security"));
        }
    }

    private void addDetailedParametersSection(XWPFTable table, JsonNode parameters) {
        XWPFTableRow headerRow = table.createRow();
        headerRow.getCell(0).setText("Параметры");
        headerRow.getCell(1).setText("");
        styleTableHeader(headerRow);

        for (JsonNode param : parameters) {
            XWPFTableRow paramRow = table.createRow();
            StringBuilder paramInfo = new StringBuilder();

            // Основная информация
            paramInfo.append("• ").append(param.path("name").asText("unnamed"))
                    .append(" (").append(param.path("in").asText("unknown")).append(")");

            if (param.path("required").asBoolean(false)) {
                paramInfo.append(" [Обязательный]");
            } else {
                paramInfo.append(" [Опциональный]");
            }

            // Тип и схема
            if (param.has("schema")) {
                String typeInfo = getDetailedSchemaInfo(param.path("schema"));
                if (!typeInfo.isEmpty()) {
                    paramInfo.append("\n  Тип: ").append(typeInfo);
                }
            }

            // Описание
            if (param.has("description") && !param.path("description").asText("").trim().isEmpty()) {
                paramInfo.append("\n  Описание: ").append(param.path("description").asText());
            }

            // Пример
            if (param.has("example")) {
                paramInfo.append("\n  Пример: ").append(param.path("example").asText());
            }

            // Разрешенные значения
            if (param.has("schema") && param.path("schema").has("enum")) {
                List<String> enumValues = new ArrayList<>();
                for (JsonNode enumValue : param.path("schema").get("enum")) {
                    enumValues.add(enumValue.asText());
                }
                paramInfo.append("\n  Допустимые значения: ").append(String.join(", ", enumValues));
            }

            paramRow.getCell(0).setText("");
            paramRow.getCell(1).setText(paramInfo.toString());
            styleTableCell(paramRow.getCell(1), false);
        }
    }

    private void addDetailedRequestBodySection(XWPFTable table, JsonNode requestBody) {
        XWPFTableRow headerRow = table.createRow();
        headerRow.getCell(0).setText("Тело запроса");
        headerRow.getCell(1).setText("");
        styleTableHeader(headerRow);

        XWPFTableRow bodyRow = table.createRow();
        StringBuilder bodyInfo = new StringBuilder();

        if (requestBody.has("description") && !requestBody.path("description").asText("").trim().isEmpty()) {
            bodyInfo.append("Описание: ").append(requestBody.path("description").asText()).append("\n");
        }

        if (requestBody.has("content")) {
            bodyInfo.append("Поддерживаемые форматы:\n");
            JsonNode content = requestBody.get("content");
            Iterator<String> contentTypes = content.fieldNames();

            while (contentTypes.hasNext()) {
                String contentType = contentTypes.next();
                bodyInfo.append("  • ").append(contentType).append("\n");

                JsonNode schema = content.get(contentType).path("schema");
                if (!schema.isMissingNode()) {
                    String schemaInfo = getDetailedSchemaInfo(schema);
                    if (!schemaInfo.isEmpty()) {
                        bodyInfo.append("    Схема: ").append(schemaInfo).append("\n");
                    }
                }

                // Примеры
                if (content.get(contentType).has("examples")) {
                    bodyInfo.append("    Примеры:\n");
                    JsonNode examples = content.get(contentType).get("examples");
                    Iterator<String> exampleNames = examples.fieldNames();
                    while (exampleNames.hasNext()) {
                        String exampleName = exampleNames.next();
                        JsonNode example = examples.get(exampleName);
                        bodyInfo.append("      - ").append(exampleName).append(": ")
                                .append(example.path("summary").asText("")).append("\n");
                    }
                }
            }
        }

        if (requestBody.has("required")) {
            bodyInfo.append("Обязательный: ").append(requestBody.path("required").asBoolean() ? "Да" : "Нет");
        }

        bodyRow.getCell(0).setText("");
        bodyRow.getCell(1).setText(bodyInfo.toString());
        styleTableCell(bodyRow.getCell(1), false);
    }

    private void addDetailedResponsesSection(XWPFTable table, JsonNode responses) {
        XWPFTableRow headerRow = table.createRow();
        headerRow.getCell(0).setText("Ответы");
        headerRow.getCell(1).setText("");
        styleTableHeader(headerRow);

        List<Map.Entry<String, JsonNode>> sortedResponses = new ArrayList<>();
        Iterator<Map.Entry<String, JsonNode>> fields = responses.fields();
        while (fields.hasNext()) {
            sortedResponses.add(fields.next());
        }

        sortedResponses.sort((e1, e2) -> {
            try {
                return Integer.compare(Integer.parseInt(e1.getKey()), Integer.parseInt(e2.getKey()));
            } catch (NumberFormatException e) {
                return e1.getKey().compareTo(e2.getKey());
            }
        });

        for (Map.Entry<String, JsonNode> entry : sortedResponses) {
            XWPFTableRow responseRow = table.createRow();
            StringBuilder responseInfo = new StringBuilder();

            String statusCode = entry.getKey();
            JsonNode response = entry.getValue();

            responseInfo.append("HTTP ").append(statusCode).append(": ");

            if (response.has("description") && !response.path("description").asText("").trim().isEmpty()) {
                responseInfo.append(response.path("description").asText()).append("\n");
            } else {
                responseInfo.append("Нет описания\n");
            }

            if (response.has("content")) {
                responseInfo.append("  Форматы ответа:\n");
                JsonNode content = response.get("content");
                Iterator<String> contentTypes = content.fieldNames();

                while (contentTypes.hasNext()) {
                    String contentType = contentTypes.next();
                    responseInfo.append("    • ").append(contentType).append("\n");

                    JsonNode schema = content.get(contentType).path("schema");
                    if (!schema.isMissingNode()) {
                        String schemaInfo = getDetailedSchemaInfo(schema);
                        if (!schemaInfo.isEmpty()) {
                            responseInfo.append("      Схема: ").append(schemaInfo).append("\n");
                        }
                    }

                    // Примеры ответа
                    if (content.get(contentType).has("example")) {
                        responseInfo.append("      Пример: ").append(
                                content.get(contentType).get("example").toString()).append("\n");
                    }
                }
            }

            // Заголовки ответа
            if (response.has("headers")) {
                responseInfo.append("  Заголовки:\n");
                JsonNode headers = response.get("headers");
                Iterator<String> headerNames = headers.fieldNames();
                while (headerNames.hasNext()) {
                    String headerName = headerNames.next();
                    JsonNode header = headers.get(headerName);
                    responseInfo.append("    • ").append(headerName).append(": ")
                            .append(header.path("description").asText("")).append("\n");
                }
            }

            responseRow.getCell(0).setText("");
            responseRow.getCell(1).setText(responseInfo.toString());
            styleTableCell(responseRow.getCell(1), false);
        }
    }

    private void addSecurityRequirementsSection(XWPFTable table, JsonNode security) {
        XWPFTableRow headerRow = table.createRow();
        headerRow.getCell(0).setText("Требования безопасности");
        headerRow.getCell(1).setText("");
        styleTableHeader(headerRow);

        for (JsonNode securityReq : security) {
            XWPFTableRow securityRow = table.createRow();
            StringBuilder securityInfo = new StringBuilder();

            Iterator<String> schemeNames = securityReq.fieldNames();
            while (schemeNames.hasNext()) {
                String schemeName = schemeNames.next();
                securityInfo.append("• ").append(schemeName);

                JsonNode scopes = securityReq.get(schemeName);
                if (scopes.size() > 0) {
                    securityInfo.append(" (области: ");
                    List<String> scopeList = new ArrayList<>();
                    for (JsonNode scope : scopes) {
                        scopeList.add(scope.asText());
                    }
                    securityInfo.append(String.join(", ", scopeList)).append(")");
                }
                securityInfo.append("\n");
            }

            securityRow.getCell(0).setText("");
            securityRow.getCell(1).setText(securityInfo.toString());
            styleTableCell(securityRow.getCell(1), false);
        }
    }

    private void addSchemasSection(XWPFDocument document, JsonNode apiSpec) {
        if (!hasSchemas(apiSpec)) {
            return;
        }

        addSectionTitle(document, "5. МОДЕЛИ ДАННЫХ", 20, "1F4E79");

        JsonNode schemas = apiSpec.get("components").get("schemas");
        List<Map.Entry<String, JsonNode>> sortedSchemas = new ArrayList<>();
        Iterator<Map.Entry<String, JsonNode>> fields = schemas.fields();
        while (fields.hasNext()) {
            sortedSchemas.add(fields.next());
        }
        sortedSchemas.sort(Comparator.comparing(Map.Entry::getKey));

        int schemaNum = 1;
        for (Map.Entry<String, JsonNode> entry : sortedSchemas) {
            addDetailedSchemaSection(document, schemaNum, entry.getKey(), entry.getValue());
            schemaNum++;
        }

        addSectionSpacing(document, 400);
    }

    private void addDetailedSchemaSection(XWPFDocument document, int schemaNum, String name, JsonNode schema) {
        addSubsectionTitle(document, "5." + schemaNum + ". " + name, 16, "2E75B6");

        // Описание
        if (schema.has("description") && !schema.path("description").asText("").trim().isEmpty()) {
            XWPFParagraph descParagraph = document.createParagraph();
            descParagraph.setAlignment(ParagraphAlignment.LEFT);
            descParagraph.setSpacingAfter(100);

            XWPFRun descRun = descParagraph.createRun();
            descRun.setText(schema.path("description").asText());
            descRun.setItalic(true);
            descRun.setFontFamily("Arial");
            descRun.setFontSize(11);
            descRun.setColor("666666");
        }

        // Информация о типе
        XWPFParagraph typeParagraph = document.createParagraph();
        typeParagraph.setAlignment(ParagraphAlignment.LEFT);
        typeParagraph.setSpacingAfter(150);

        XWPFRun typeRun = typeParagraph.createRun();
        typeRun.setText("Тип данных: ");
        typeRun.setBold(true);
        typeRun.setFontFamily("Arial");
        typeRun.setFontSize(11);

        String typeInfo = getDetailedSchemaInfo(schema);
        typeRun.setBold(false);
        typeRun.setText(typeInfo);

        // Свойства
        if (schema.has("properties") && schema.get("properties").size() > 0) {
            createDetailedPropertiesTable(document, schema);
        }

        // Enum values
        if (schema.has("enum") && schema.get("enum").size() > 0) {
            XWPFParagraph enumParagraph = document.createParagraph();
            enumParagraph.setAlignment(ParagraphAlignment.LEFT);
            enumParagraph.setSpacingAfter(100);

            XWPFRun enumRun = enumParagraph.createRun();
            enumRun.setText("Допустимые значения: ");
            enumRun.setBold(true);
            enumRun.setFontFamily("Arial");
            enumRun.setFontSize(10);

            List<String> enumValues = new ArrayList<>();
            for (JsonNode enumValue : schema.get("enum")) {
                enumValues.add(enumValue.asText());
            }
            enumRun.setBold(false);
            enumRun.setText(String.join(", ", enumValues));
            enumRun.setColor("8B4513");
        }

        // Пример
        if (schema.has("example")) {
            XWPFParagraph exampleParagraph = document.createParagraph();
            exampleParagraph.setAlignment(ParagraphAlignment.LEFT);
            exampleParagraph.setSpacingAfter(100);
            exampleParagraph.setIndentationLeft(200);

            XWPFRun exampleRun = exampleParagraph.createRun();
            exampleRun.setText("Пример: ");
            exampleRun.setBold(true);
            exampleRun.setFontFamily("Courier New");
            exampleRun.setFontSize(9);

            exampleRun.setBold(false);
            exampleRun.setText(schema.path("example").toString());
            exampleRun.setColor("2E8B57");
        }

        addSectionSpacing(document, 300);
    }

    private void createDetailedPropertiesTable(XWPFDocument document, JsonNode schema) {
        XWPFTable table = document.createTable(1, 5);
        table.setWidth("100%");
        setTableColumnWidths(table, new int[]{20, 15, 10, 15, 40});

        // Заголовки
        XWPFTableRow headerRow = table.getRow(0);
        headerRow.getCell(0).setText("Поле");
        headerRow.getCell(1).setText("Тип");
        headerRow.getCell(2).setText("Обяз.");
        headerRow.getCell(3).setText("Формат");
        headerRow.getCell(4).setText("Описание");
        styleTableHeader(headerRow);

        Set<String> requiredFields = new HashSet<>();
        if (schema.has("required")) {
            for (JsonNode requiredField : schema.get("required")) {
                requiredFields.add(requiredField.asText());
            }
        }

        JsonNode properties = schema.get("properties");
        List<Map.Entry<String, JsonNode>> sortedProperties = new ArrayList<>();
        Iterator<Map.Entry<String, JsonNode>> fields = properties.fields();
        while (fields.hasNext()) {
            sortedProperties.add(fields.next());
        }
        sortedProperties.sort(Comparator.comparing(Map.Entry::getKey));

        for (Map.Entry<String, JsonNode> property : sortedProperties) {
            String fieldName = property.getKey();
            JsonNode fieldSchema = property.getValue();

            XWPFTableRow row = table.createRow();
            row.getCell(0).setText(fieldName);
            row.getCell(1).setText(fieldSchema.path("type").asText("object"));
            row.getCell(2).setText(requiredFields.contains(fieldName) ? "Да" : "Нет");
            row.getCell(3).setText(fieldSchema.path("format").asText(""));

            String description = fieldSchema.has("description") ?
                    fieldSchema.get("description").asText() : "";
            row.getCell(4).setText(description);

            // Подсветка обязательных полей
            if (requiredFields.contains(fieldName)) {
                row.getCell(0).setColor("FFF2CC");
            }
        }
    }

    private void addSecuritySchemesSection(XWPFDocument document, JsonNode apiSpec) {
        if (!hasSecuritySchemes(apiSpec)) {
            return;
        }

        addSectionTitle(document, "6. БЕЗОПАСНОСТЬ", 20, "1F4E79");

        JsonNode securitySchemes = apiSpec.get("components").get("securitySchemes");
        Iterator<Map.Entry<String, JsonNode>> schemes = securitySchemes.fields();
        int schemeNum = 1;
        while (schemes.hasNext()) {
            Map.Entry<String, JsonNode> scheme = schemes.next();
            addDetailedSecurityScheme(document, schemeNum, scheme.getKey(), scheme.getValue());
            schemeNum++;
        }

        addSectionSpacing(document, 400);
    }

    private void addDetailedSecurityScheme(XWPFDocument document, int schemeNum, String name, JsonNode scheme) {
        addSubsectionTitle(document, "6." + schemeNum + ". " + name, 16, "2E75B6");

        XWPFTable table = document.createTable(1, 2);
        table.setWidth("100%");
        setTableColumnWidths(table, new int[]{30, 70});

        addTableRow(table, "Тип", scheme.path("type").asText("N/A"), 0);

        if (scheme.has("description") && !scheme.path("description").asText("").trim().isEmpty()) {
            addTableRow(table, "Описание", scheme.path("description").asText(), 1);
        }

        if (scheme.has("scheme") && !scheme.path("scheme").asText("").trim().isEmpty()) {
            addTableRow(table, "Схема", scheme.path("scheme").asText(), 2);
        }

        if (scheme.has("bearerFormat") && !scheme.path("bearerFormat").asText("").trim().isEmpty()) {
            addTableRow(table, "Формат Bearer", scheme.path("bearerFormat").asText(), 3);
        }

        if (scheme.has("in") && !scheme.path("in").asText("").trim().isEmpty()) {
            addTableRow(table, "Расположение", scheme.path("in").asText(), 4);
        }

        if (scheme.has("name") && !scheme.path("name").asText("").trim().isEmpty()) {
            addTableRow(table, "Имя", scheme.path("name").asText(), 5);
        }

        // Потоки OAuth
        if (scheme.has("flows")) {
            JsonNode flows = scheme.get("flows");
            addTableRow(table, "Потоки OAuth", "Определены следующие потоки:", 6);

            Iterator<String> flowTypes = flows.fieldNames();
            while (flowTypes.hasNext()) {
                String flowType = flowTypes.next();
                JsonNode flow = flows.get(flowType);
                XWPFTableRow flowRow = table.createRow();
                flowRow.getCell(0).setText("");
                flowRow.getCell(1).setText("  • " + flowType + ": " +
                        flow.path("authorizationUrl").asText("") +
                        flow.path("tokenUrl").asText(""));
                styleTableCell(flowRow.getCell(1), false);
            }
        }

        addSectionSpacing(document, 300);
    }

    // Вспомогательные методы
    private void addSectionTitle(XWPFDocument document, String title, int fontSize, String color) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        paragraph.setSpacingBefore(600);
        paragraph.setSpacingAfter(200);

        XWPFRun run = paragraph.createRun();
        run.setText(title);
        run.setBold(true);
        run.setFontSize(fontSize);
        run.setFontFamily("Arial");
        run.setColor(color);
    }

    private void addSubsectionTitle(XWPFDocument document, String title, int fontSize, String color) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        paragraph.setSpacingBefore(400);
        paragraph.setSpacingAfter(150);

        XWPFRun run = paragraph.createRun();
        run.setText(title);
        run.setBold(true);
        run.setFontSize(fontSize);
        run.setFontFamily("Arial");
        run.setColor(color);
    }

    private void setTableColumnWidths(XWPFTable table, int[] percentages) {
        if (table.getNumberOfRows() == 0) return;

        XWPFTableRow row = table.getRow(0);
        int totalWidth = 10000; // Total table width in twips

        for (int i = 0; i < percentages.length && i < row.getTableCells().size(); i++) {
            CTTblWidth width = CTTblWidth.Factory.newInstance();
            width.setW(BigInteger.valueOf(totalWidth * percentages[i] / 100));
            width.setType(STTblWidth.PCT);
            row.getCell(i).getCTTc().addNewTcPr().setTcW(width);
        }
    }

    private void addTableRow(XWPFTable table, String header, String value, int rowIndex) {
        XWPFTableRow row;
        if (rowIndex < table.getNumberOfRows()) {
            row = table.getRow(rowIndex);
        } else {
            row = table.createRow();
        }

        row.getCell(0).setText(header);
        row.getCell(1).setText(value != null ? value : "N/A");

        styleTableCell(row.getCell(0), true);
        styleTableCell(row.getCell(1), false);
    }

    private String getDetailedSchemaInfo(JsonNode schema) {
        if (schema.has("$ref")) {
            return extractSchemaName(schema.path("$ref").asText()) + " (ссылка)";
        }

        StringBuilder info = new StringBuilder();

        if (schema.has("type")) {
            info.append(schema.path("type").asText());
        }

        if (schema.has("format")) {
            info.append(" (").append(schema.path("format").asText()).append(")");
        }

        if (schema.has("items")) {
            info.append(" массива [").append(getDetailedSchemaInfo(schema.get("items"))).append("]");
        }

        // Дополнительные свойства
        if (schema.has("minimum")) {
            info.append(", мин: ").append(schema.path("minimum").asText());
        }
        if (schema.has("maximum")) {
            info.append(", макс: ").append(schema.path("maximum").asText());
        }
        if (schema.has("pattern")) {
            info.append(", паттерн: ").append(schema.path("pattern").asText());
        }
        if (schema.has("minLength")) {
            info.append(", мин. длина: ").append(schema.path("minLength").asText());
        }
        if (schema.has("maxLength")) {
            info.append(", макс. длина: ").append(schema.path("maxLength").asText());
        }

        return info.length() > 0 ? info.toString() : "object";
    }

    private String extractSchemaName(String ref) {
        if (ref == null) return "";
        Pattern pattern = Pattern.compile(".*/([^/]+)$");
        java.util.regex.Matcher matcher = pattern.matcher(ref);
        return matcher.find() ? matcher.group(1) : ref;
    }

    private String getMethodColor(String method) {
        switch (method.toUpperCase()) {
            case "GET": return "2E8B57";
            case "POST": return "1E90FF";
            case "PUT": return "FF8C00";
            case "DELETE": return "DC143C";
            case "PATCH": return "FFD700";
            case "HEAD": return "9932CC";
            case "OPTIONS": return "808080";
            default: return "696969";
        }
    }

    private void styleTableHeader(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            cell.setColor("4472C4");
            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                paragraph.setSpacingAfter(0);
                paragraph.setSpacingBefore(0);
                for (XWPFRun run : paragraph.getRuns()) {
                    run.setBold(true);
                    run.setFontFamily("Arial");
                    run.setFontSize(10);
                    run.setColor("FFFFFF");
                }
            }
        }
    }

    private void styleTableCell(XWPFTableCell cell, boolean isHeader) {
        if (isHeader) {
            cell.setColor("D9E2F3");
        }

        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            paragraph.setSpacingAfter(0);
            paragraph.setSpacingBefore(0);
            for (XWPFRun run : paragraph.getRuns()) {
                run.setFontFamily("Arial");
                run.setFontSize(9);
                run.setColor("333333");
                if (isHeader) {
                    run.setBold(true);
                }
            }
        }
    }

    private void addEmptyEndpointsMessage(XWPFDocument document) {
        XWPFParagraph emptyMsg = document.createParagraph();
        emptyMsg.setAlignment(ParagraphAlignment.LEFT);
        emptyMsg.setSpacingBefore(200);

        XWPFRun emptyRun = emptyMsg.createRun();
        emptyRun.setText("В спецификации API не определены endpoints.");
        emptyRun.setFontSize(11);
        emptyRun.setFontFamily("Arial");
        emptyRun.setColor("666666");
        emptyRun.setItalic(true);
    }

    private void addPageBreak(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.addBreak(BreakType.PAGE);
    }

    private void addSectionSpacing(XWPFDocument document, int spacing) {
        XWPFParagraph spacer = document.createParagraph();
        spacer.setSpacingAfter(spacing);
    }

    // Проверки наличия разделов
    private boolean hasServers(JsonNode apiSpec) {
        return apiSpec.has("servers") && apiSpec.get("servers").size() > 0;
    }

    private boolean hasTags(JsonNode apiSpec) {
        return apiSpec.has("tags") && apiSpec.get("tags").size() > 0;
    }

    private boolean hasSchemas(JsonNode apiSpec) {
        return apiSpec.has("components") &&
                apiSpec.get("components").has("schemas") &&
                apiSpec.get("components").get("schemas").size() > 0;
    }

    private boolean hasSecuritySchemes(JsonNode apiSpec) {
        return apiSpec.has("components") &&
                apiSpec.get("components").has("securitySchemes") &&
                apiSpec.get("components").get("securitySchemes").size() > 0;
    }

    // Вспомогательный класс
    private static class EndpointInfo {
        private final String path;
        private final String method;
        private final JsonNode operation;

        public EndpointInfo(String path, String method, JsonNode operation) {
            this.path = path;
            this.method = method;
            this.operation = operation;
        }

        public String getPath() { return path; }
        public String getMethod() { return method; }
        public JsonNode getOperation() { return operation; }
    }
}