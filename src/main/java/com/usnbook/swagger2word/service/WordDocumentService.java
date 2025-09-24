package com.usnbook.swagger2word.service;

import com.usnbook.swagger2word.model.OpenApiSpec;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import org.slf4j.Logger;

@Service
public class WordDocumentService {

    @Value("${app.output-directory:./generated-docs}")
    private String outputDirectory;

    private static final Logger logger = LoggerFactory.getLogger(ApiDocsService.class);

    public String generateWordDocument(OpenApiSpec apiSpec) throws Exception {
        if (apiSpec == null) {
            throw new IllegalArgumentException("API спецификация не может быть null");
        }
        if (apiSpec.getInfo() == null) {
            throw new IllegalArgumentException("Информация об API отсутствует");
        }

        try (XWPFDocument document = new XWPFDocument()) {
            addTitlePage(document, apiSpec);
            addGeneralInfo(document, apiSpec);
            addServersSection(document, apiSpec.getServers());
            addTagsSection(document, apiSpec.getTags());
            addEndpointsByTags(document, apiSpec);
            addSchemasSection(document, apiSpec.getComponents());

            String fileName = generateFileName(apiSpec.getInfo().getTitle());
            Path filePath = Paths.get(fileName);
            Files.createDirectories(filePath.getParent());

            try (FileOutputStream out = new FileOutputStream(filePath.toFile())) {
                document.write(out);
            }

            return fileName;
        }
    }


    private void addTitlePage(XWPFDocument document, OpenApiSpec apiSpec) {
        XWPFParagraph titleParagraph = document.createParagraph();
        titleParagraph.setAlignment(ParagraphAlignment.CENTER);
        titleParagraph.setSpacingBefore(600);
        titleParagraph.setSpacingAfter(400);

        XWPFRun titleRun = titleParagraph.createRun();
        titleRun.setText("ДОКУМЕНТАЦИЯ API");
        titleRun.setBold(true);
        titleRun.setFontSize(24);
        titleRun.setFontFamily("Times New Roman");
        titleRun.setColor("000000");
        titleRun.addBreak();

        XWPFParagraph apiTitleParagraph = document.createParagraph();
        apiTitleParagraph.setAlignment(ParagraphAlignment.CENTER);
        apiTitleParagraph.setSpacingAfter(300);

        XWPFRun apiTitleRun = apiTitleParagraph.createRun();
        apiTitleRun.setText(apiSpec.getInfo().getTitle());
        apiTitleRun.setBold(true);
        apiTitleRun.setFontSize(18);
        apiTitleRun.setFontFamily("Times New Roman");
        apiTitleRun.setColor("000000");
        apiTitleRun.addBreak();

        if (apiSpec.getInfo().getDescription() != null) {
            XWPFParagraph descParagraph = document.createParagraph();
            descParagraph.setAlignment(ParagraphAlignment.CENTER);
            descParagraph.setSpacingAfter(400);
            descParagraph.setIndentationLeft(720);
            descParagraph.setIndentationRight(720);

            XWPFRun descRun = descParagraph.createRun();
            descRun.setText(apiSpec.getInfo().getDescription());
            descRun.setFontSize(12);
            descRun.setFontFamily("Times New Roman");
            descRun.setColor("000000");
        }

        if (apiSpec.getInfo().getContact() != null) {
            addContactInfo(document, apiSpec.getInfo().getContact());
        }

        if (apiSpec.getInfo().getLicense() != null) {
            addLicenseInfo(document, apiSpec.getInfo().getLicense());
        }

        XWPFParagraph metaParagraph = document.createParagraph();
        metaParagraph.setAlignment(ParagraphAlignment.CENTER);
        metaParagraph.setSpacingAfter(400);

        XWPFRun metaRun = metaParagraph.createRun();
        metaRun.setText("Версия API: " + (apiSpec.getInfo().getVersion() != null ? apiSpec.getInfo().getVersion() : "N/A"));
        metaRun.setFontSize(10);
        metaRun.setFontFamily("Times New Roman");
        metaRun.setColor("666666");
        metaRun.addBreak();
        metaRun.setText("OpenAPI Version: " + (apiSpec.getOpenapi() != null ? apiSpec.getOpenapi() : "N/A"));
        metaRun.addBreak();
        metaRun.setText("Сгенерировано: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm:ss")));

        addPageBreak(document);
    }

    private void addContactInfo(XWPFDocument document, OpenApiSpec.Contact contact) {
        XWPFParagraph contactTitle = document.createParagraph();
        contactTitle.setAlignment(ParagraphAlignment.CENTER);
        contactTitle.setSpacingBefore(200);
        contactTitle.setSpacingAfter(100);

        XWPFRun contactTitleRun = contactTitle.createRun();
        contactTitleRun.setText("КОНТАКТНАЯ ИНФОРМАЦИЯ");
        contactTitleRun.setBold(true);
        contactTitleRun.setFontSize(12);
        contactTitleRun.setFontFamily("Times New Roman");
        contactTitleRun.setColor("000000");
        contactTitleRun.addBreak();

        XWPFParagraph contactParagraph = document.createParagraph();
        contactParagraph.setAlignment(ParagraphAlignment.CENTER);
        contactParagraph.setSpacingAfter(200);

        XWPFRun contactRun = contactParagraph.createRun();
        if (contact.getName() != null) {
            contactRun.setText("Разработчик: " + contact.getName());
            contactRun.addBreak();
        }
        if (contact.getEmail() != null) {
            contactRun.setText("Email: " + contact.getEmail());
            contactRun.addBreak();
        }
        if (contact.getUrl() != null) {
            contactRun.setText("Сайт: " + contact.getUrl());
        }

        contactRun.setFontSize(10);
        contactRun.setFontFamily("Times New Roman");
        contactRun.setColor("000000");
    }

    private void addLicenseInfo(XWPFDocument document, OpenApiSpec.License license) {
        XWPFParagraph licenseParagraph = document.createParagraph();
        licenseParagraph.setAlignment(ParagraphAlignment.CENTER);
        licenseParagraph.setSpacingAfter(200);

        XWPFRun licenseRun = licenseParagraph.createRun();
        licenseRun.setText("Лицензия: " + license.getName());
        if (license.getUrl() != null) {
            licenseRun.addBreak();
            licenseRun.setText("URL: " + license.getUrl());
        }

        licenseRun.setFontSize(10);
        licenseRun.setFontFamily("Times New Roman");
        licenseRun.setColor("000000");
    }

    private void addGeneralInfo(XWPFDocument document, OpenApiSpec apiSpec) {
        XWPFParagraph sectionTitle = document.createParagraph();
        sectionTitle.setStyle("Heading1");
        sectionTitle.setSpacingBefore(600);
        sectionTitle.setSpacingAfter(200);

        XWPFRun titleRun = sectionTitle.createRun();
        titleRun.setText("1. ОБЩАЯ ИНФОРМАЦИЯ");
        titleRun.setBold(true);
        titleRun.setFontSize(16);
        titleRun.setFontFamily("Times New Roman");
        titleRun.setColor("000000");

        XWPFParagraph infoParagraph = document.createParagraph();
        infoParagraph.setSpacingAfter(300);
        infoParagraph.setIndentationLeft(720);

        XWPFRun infoRun = infoParagraph.createRun();
        infoRun.setFontFamily("Times New Roman");
        infoRun.setFontSize(11);
        infoRun.setColor("000000");

        infoRun.setText("Название: ");
        infoRun.setBold(true);
        infoRun.setText(apiSpec.getInfo().getTitle());
        infoRun.addBreak();
        infoRun.setBold(false);

        infoRun.setText("Версия: ");
        infoRun.setBold(true);
        infoRun.setText(apiSpec.getInfo().getVersion() != null ? apiSpec.getInfo().getVersion() : "N/A");
        infoRun.addBreak();
        infoRun.setBold(false);

        infoRun.setText("OpenAPI: ");
        infoRun.setBold(true);
        infoRun.setText(apiSpec.getOpenapi() != null ? apiSpec.getOpenapi() : "N/A");
        infoRun.addBreak();
        infoRun.setBold(false);

        if (apiSpec.getInfo().getDescription() != null) {
            infoRun.addBreak();
            infoRun.setText("Описание: ");
            infoRun.addBreak();
            infoRun.setText("   " + apiSpec.getInfo().getDescription());
        }
    }

    private void addServersSection(XWPFDocument document, List<OpenApiSpec.Server> servers) {
        if (servers != null && !servers.isEmpty()) {
            XWPFParagraph serversTitle = document.createParagraph();
            serversTitle.setStyle("Heading2");
            serversTitle.setSpacingBefore(600);
            serversTitle.setSpacingAfter(200);

            XWPFRun serversTitleRun = serversTitle.createRun();
            serversTitleRun.setText("2. СЕРВЕРЫ");
            serversTitleRun.setBold(true);
            serversTitleRun.setFontSize(14);
            serversTitleRun.setFontFamily("Times New Roman");
            serversTitleRun.setColor("000000");

            int serverNum = 1;
            for (OpenApiSpec.Server server : servers) {
                XWPFParagraph serverParagraph = document.createParagraph();
                serverParagraph.setSpacingAfter(150);
                serverParagraph.setIndentationLeft(720);

                XWPFRun serverRun = serverParagraph.createRun();
                serverRun.setFontFamily("Times New Roman");
                serverRun.setFontSize(11);
                serverRun.setColor("000000");

                serverRun.setText(serverNum + ". ");
                serverRun.setBold(true);
                serverRun.setText(server.getUrl());
                serverRun.setBold(false);
                serverRun.addBreak();

                if (server.getDescription() != null) {
                    serverRun.setText("   Описание: " + server.getDescription());
                    serverRun.setItalic(true);
                }

                serverNum++;
            }

            addSectionSpacing(document);
        }
    }

    private void addTagsSection(XWPFDocument document, List<OpenApiSpec.Tag> tags) {
        if (tags != null && !tags.isEmpty()) {
            XWPFParagraph tagsTitle = document.createParagraph();
            tagsTitle.setStyle("Heading2");
            tagsTitle.setSpacingBefore(600);
            tagsTitle.setSpacingAfter(200);

            XWPFRun tagsTitleRun = tagsTitle.createRun();
            tagsTitleRun.setText("3. ГРУППЫ API");
            tagsTitleRun.setBold(true);
            tagsTitleRun.setFontSize(14);
            tagsTitleRun.setFontFamily("Times New Roman");
            tagsTitleRun.setColor("000000");

            int tagNum = 1;
            for (OpenApiSpec.Tag tag : tags) {
                XWPFParagraph tagParagraph = document.createParagraph();
                tagParagraph.setSpacingAfter(150);
                tagParagraph.setIndentationLeft(720);

                XWPFRun tagRun = tagParagraph.createRun();
                tagRun.setFontFamily("Times New Roman");
                tagRun.setFontSize(11);
                tagRun.setColor("000000");

                tagRun.setText(tagNum + ". ");
                tagRun.setBold(true);
                tagRun.setText(tag.getName());
                tagRun.setBold(false);
                tagRun.addBreak();

                if (tag.getDescription() != null) {
                    tagRun.setText("   " + tag.getDescription());
                    tagRun.setItalic(true);
                }

                tagNum++;
            }

            addSectionSpacing(document);
        }
    }

    private void addEndpointsByTags(XWPFDocument document, OpenApiSpec apiSpec) {
        XWPFParagraph endpointsTitle = document.createParagraph();
        endpointsTitle.setStyle("Heading1");
        endpointsTitle.setSpacingBefore(600);
        endpointsTitle.setSpacingAfter(200);

        XWPFRun endpointsTitleRun = endpointsTitle.createRun();
        endpointsTitleRun.setText("4. ENDPOINTS");
        endpointsTitleRun.setBold(true);
        endpointsTitleRun.setFontSize(16);
        endpointsTitleRun.setFontFamily("Times New Roman");
        endpointsTitleRun.setColor("000000");

        Map<String, OpenApiSpec.Path> allPaths = getAllPaths(apiSpec);

        if (allPaths == null || allPaths.isEmpty()) {
            addEmptyEndpointsMessage(document);
            addSectionSpacing(document);
            return;
        }

        // === ДИАГНОСТИКА ===
        DiagnosticInfo diagnostics = analyzeApiStructure(allPaths);
        addDiagnosticInfo(document, diagnostics);

        // === ГРУППИРОВКА И ОТОБРАЖЕНИЕ ===
        Map<String, List<EndpointOperation>> groupedOperations = groupOperationsByTags(allPaths, diagnostics);

        if (!groupedOperations.isEmpty() && hasValidOperations(groupedOperations)) {
            displayGroupedEndpoints(document, groupedOperations, diagnostics);
        } else {
            displayAllEndpointsFallback(document, allPaths, diagnostics);
        }

        addSectionSpacing(document);
    }


    private Map<String, OpenApiSpec.Path> getAllPaths(OpenApiSpec apiSpec) {
        Map<String, OpenApiSpec.Path> allPaths = new LinkedHashMap<>();

        Object pathsObject = apiSpec.getPaths();

        if (pathsObject == null) {
            logger.warn("apiSpec.getPaths() returned null");
            return allPaths;
        }

        logger.debug("Paths object type: {}", pathsObject.getClass().getSimpleName());


        if (pathsObject instanceof Map) {
            @SuppressWarnings("unchecked")
            Map<String, OpenApiSpec.Path> pathsMap = (Map<String, OpenApiSpec.Path>) pathsObject;
            if (!pathsMap.isEmpty()) {
                allPaths.putAll(pathsMap);
                logger.debug("Loaded {} paths from Map", pathsMap.size());
            }
        }

        else {
            try {
                java.lang.reflect.Method getPathsMethod = pathsObject.getClass().getMethod("getPaths");
                @SuppressWarnings("unchecked")
                Map<String, OpenApiSpec.Path> pathsMap = (Map<String, OpenApiSpec.Path>) getPathsMethod.invoke(pathsObject);

                if (pathsMap != null && !pathsMap.isEmpty()) {
                    allPaths.putAll(pathsMap);
                    logger.debug("Loaded {} paths from Paths object", pathsMap.size());
                }
            } catch (Exception e) {
                logger.warn("Could not extract paths via reflection: {}", e.getMessage());
            }
        }

        if (allPaths.isEmpty()) {
            logger.warn("No paths found in API specification");
        }

        return allPaths;
    }

    private Map<String, OpenApiSpec.Operation> extractOperationsFromPath(OpenApiSpec.Path pathItem) {
        Map<String, OpenApiSpec.Operation> operations = new LinkedHashMap<>();

        if (pathItem == null) return null;

        try {
            Map<String, OpenApiSpec.Operation> allOps = pathItem.getAllOperations();
            if (allOps != null && !allOps.isEmpty()) {
                operations.putAll(allOps);
            }
        } catch (Exception e) {

        }


        if (operations.isEmpty()) {
            try {
                if (pathItem.getOperations() != null && !pathItem.getOperations().isEmpty()) {
                    operations.putAll(pathItem.getOperations());
                }
            } catch (Exception e) {

            }
        }


        if (operations.isEmpty()) {
            try {
                if (pathItem.getGetOperation() != null) operations.put("get", pathItem.getGetOperation());
                if (pathItem.getPostOperation() != null) operations.put("post", pathItem.getPostOperation());
                if (pathItem.getPutOperation() != null) operations.put("put", pathItem.getPutOperation());
                if (pathItem.getDeleteOperation() != null) operations.put("delete", pathItem.getDeleteOperation());
                if (pathItem.getPatchOperation() != null) operations.put("patch", pathItem.getPatchOperation());
            } catch (Exception e) {

            }
        }

        return operations.isEmpty() ? null : operations;
    }

    // === АНАЛИЗ СТРУКТУРЫ API ===
    private DiagnosticInfo analyzeApiStructure(Map<String, OpenApiSpec.Path> allPaths) {
        DiagnosticInfo info = new DiagnosticInfo();

        info.totalPaths = allPaths != null ? allPaths.size() : 0;

        if (allPaths != null) {
            for (Map.Entry<String, OpenApiSpec.Path> pathEntry : allPaths.entrySet()) {
                String path = pathEntry.getKey();
                OpenApiSpec.Path pathItem = pathEntry.getValue();

                PathInfo pathInfo = new PathInfo();
                pathInfo.path = path;

                Map<String, OpenApiSpec.Operation> operations = extractOperationsFromPath(pathItem);
                pathInfo.operations = operations != null ? operations.size() : 0;

                if (operations != null) {
                    for (Map.Entry<String, OpenApiSpec.Operation> opEntry : operations.entrySet()) {
                        OpenApiSpec.Operation operation = opEntry.getValue();
                        if (operation != null) {
                            OperationInfo opInfo = new OperationInfo();
                            opInfo.method = opEntry.getKey() != null ? opEntry.getKey().toUpperCase() : "UNKNOWN";
                            opInfo.operationId = operation.getOperationId();
                            opInfo.summary = operation.getSummary();

                            List<String> tags = operation.getTags();
                            if (tags != null && !tags.isEmpty()) {
                                opInfo.tags.addAll(tags);
                                info.totalTaggedOperations++;
                            } else {
                                info.totalUntaggedOperations++;
                            }

                            pathInfo.operationsInfo.add(opInfo);
                        }
                    }
                }

                info.pathInfos.add(pathInfo);
            }
        }

        return info;
    }

    // === ГРУППИРОВКА ОПЕРАЦИЙ ПО ТЕГАМ ===
    private Map<String, List<EndpointOperation>> groupOperationsByTags(Map<String, OpenApiSpec.Path> allPaths,
                                                                       DiagnosticInfo diagnostics) {
        Map<String, List<EndpointOperation>> grouped = new LinkedHashMap<>();

        for (PathInfo pathInfo : diagnostics.pathInfos) {
            OpenApiSpec.Path pathItem = allPaths.get(pathInfo.path);
            if (pathItem == null) continue;

            Map<String, OpenApiSpec.Operation> operations = extractOperationsFromPath(pathItem);
            if (operations == null || operations.isEmpty()) continue;

            for (Map.Entry<String, OpenApiSpec.Operation> opEntry : operations.entrySet()) {
                OpenApiSpec.Operation operation = opEntry.getValue();
                if (operation == null) continue;

                EndpointOperation endpointOp = new EndpointOperation();
                endpointOp.path = pathInfo.path;
                endpointOp.method = opEntry.getKey() != null ? opEntry.getKey().toUpperCase() : "UNKNOWN";
                endpointOp.operation = operation;

                List<String> opTags = operation.getTags();

                if (opTags != null && !opTags.isEmpty()) {
                    for (String tag : opTags) {
                        if (tag != null && !tag.trim().isEmpty()) {
                            grouped.computeIfAbsent(tag.trim(), k -> new ArrayList<>()).add(endpointOp);
                        }
                    }
                } else {

                    String inferredTag = inferTagFromOperation(operation, diagnostics.specTags);
                    if (inferredTag != null && !inferredTag.trim().isEmpty()) {
                        grouped.computeIfAbsent(inferredTag.trim(), k -> new ArrayList<>()).add(endpointOp);
                    } else {
                        grouped.computeIfAbsent("Не классифицировано", k -> new ArrayList<>()).add(endpointOp);
                    }
                }
            }
        }

        return grouped;
    }

    // === ПРОВЕРКА НА ВАЛИДНЫЕ ОПЕРАЦИИ ===
    private boolean hasValidOperations(Map<String, List<EndpointOperation>> groupedOperations) {
        for (List<EndpointOperation> operations : groupedOperations.values()) {
            if (operations != null && !operations.isEmpty()) {
                for (EndpointOperation op : operations) {
                    if (op != null && op.operation != null &&
                            (op.operation.getOperationId() != null ||
                                    op.operation.getSummary() != null ||
                                    op.operation.getParameters() != null ||
                                    op.operation.getResponses() != null)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    // === ОПРЕДЕЛЕНИЕ ТЕГА ПО OPERATIONID ===
    private String inferTagFromOperation(OpenApiSpec.Operation operation, List<String> availableTags) {
        String operationId = operation.getOperationId();
        String summary = operation.getSummary();

        if (operationId != null && !operationId.trim().isEmpty()) {
            String opIdClean = operationId.toLowerCase().replaceAll("[^a-zа-я0-9]", "");

            for (String tag : availableTags) {
                if (tag == null) continue;
                String tagClean = tag.toLowerCase().replaceAll("[^a-zа-я0-9]", "");
                if (!tagClean.isEmpty() && (opIdClean.contains(tagClean) || tagClean.contains(opIdClean))) {
                    return tag;
                }
            }
        }

        if (summary != null && !summary.trim().isEmpty()) {
            String summaryLower = summary.toLowerCase();
            for (String tag : availableTags) {
                if (tag != null && summaryLower.contains(tag.toLowerCase())) {
                    return tag;
                }
            }
        }

        return null;
    }


    private void addDiagnosticInfo(XWPFDocument document, DiagnosticInfo diagnostics) {
        XWPFParagraph diagTitle = document.createParagraph();
        diagTitle.setSpacingBefore(100);
        diagTitle.setSpacingAfter(50);
        diagTitle.setIndentationLeft(360);

        XWPFRun diagTitleRun = diagTitle.createRun();
        diagTitleRun.setText("=== ДИАГНОСТИКА API ===");
        diagTitleRun.setBold(true);
        diagTitleRun.setFontSize(12);
        diagTitleRun.setFontFamily("Courier New");
        diagTitleRun.setColor("FF0000");
        diagTitleRun.addBreak();

        XWPFParagraph diagContent = document.createParagraph();
        diagContent.setSpacingAfter(100);
        diagContent.setIndentationLeft(720);

        XWPFRun diagRun = diagContent.createRun();
        diagRun.setFontFamily("Courier New");
        diagRun.setFontSize(9);
        diagRun.setColor("000000");

        diagRun.setText("СТРУКТУРА PATHS:");
        diagRun.addBreak();
        diagRun.setText("  Всего путей: " + diagnostics.totalPaths);
        diagRun.addBreak();

        int totalOps = diagnostics.totalTaggedOperations + diagnostics.totalUntaggedOperations;
        diagRun.setText("  Всего операций: " + totalOps);
        diagRun.addBreak();
        diagRun.setText("  С тегами: " + diagnostics.totalTaggedOperations);
        diagRun.addBreak();
        diagRun.setText("  Без тегов: " + diagnostics.totalUntaggedOperations);
        diagRun.addBreak();

        diagRun.addBreak();
        diagRun.setText("ТЕГИ ИЗ СПЕЦИФИКАЦИИ (" + diagnostics.specTags.size() + "):");
        for (int i = 0; i < Math.min(5, diagnostics.specTags.size()); i++) {
            diagRun.setText("  " + (i + 1) + ". " + diagnostics.specTags.get(i));
            diagRun.addBreak();
        }
        if (diagnostics.specTags.size() > 5) {
            diagRun.setText("  ... и еще " + (diagnostics.specTags.size() - 5));
            diagRun.addBreak();
        }

        diagRun.addBreak();
        diagRun.setText("ПЕРВЫЕ 5 ПУТЕЙ:");
        for (int i = 0; i < Math.min(5, diagnostics.pathInfos.size()); i++) {
            PathInfo pathInfo = diagnostics.pathInfos.get(i);
            diagRun.setText("  " + (i + 1) + ". " + pathInfo.path + " (" + pathInfo.operations + " оп.)");
            diagRun.addBreak();
        }
        if (diagnostics.pathInfos.size() > 5) {
            diagRun.setText("  ... и еще " + (diagnostics.pathInfos.size() - 5) + " путей");
            diagRun.addBreak();
        }

        addSectionSpacing(document, 200);
    }

    private void displayGroupedEndpoints(XWPFDocument document,
                                         Map<String, List<EndpointOperation>> groupedOperations,
                                         DiagnosticInfo diagnostics) {
        List<String> sortedGroups = new ArrayList<>(groupedOperations.keySet());
        Collections.sort(sortedGroups, (a, b) -> {
            if ("Не классифицировано".equals(a)) return 1;
            if ("Не классифицировано".equals(b)) return -1;
            return a.compareToIgnoreCase(b);
        });

        int groupNum = 1;
        for (String groupName : sortedGroups) {
            List<EndpointOperation> operations = groupedOperations.get(groupName);
            if (operations != null && !operations.isEmpty()) {
                addGroupSection(document, groupNum, groupName, operations);
                groupNum++;
            }
        }
    }

    private void displayAllEndpointsFallback(XWPFDocument document, Map<String, OpenApiSpec.Path> allPaths,
                                             DiagnosticInfo diagnostics) {
        XWPFParagraph fallbackTitle = document.createParagraph();
        fallbackTitle.setStyle("Heading2");
        fallbackTitle.setSpacingBefore(200);
        fallbackTitle.setSpacingAfter(100);

        XWPFRun fallbackTitleRun = fallbackTitle.createRun();
        fallbackTitleRun.setText("4.1. ВСЕ ENDPOINTS (без группировки)");
        fallbackTitleRun.setBold(true);
        fallbackTitleRun.setFontSize(14);
        fallbackTitleRun.setFontFamily("Times New Roman");
        fallbackTitleRun.setColor("000000");

        XWPFParagraph fallbackNote = document.createParagraph();
        fallbackNote.setSpacingAfter(100);
        fallbackNote.setIndentationLeft(720);

        XWPFRun fallbackNoteRun = fallbackNote.createRun();
        fallbackNoteRun.setText("Примечание: Автоматическая группировка по тегам недоступна. Показаны все найденные endpoints.");
        fallbackNoteRun.setFontSize(10);
        fallbackNoteRun.setFontFamily("Times New Roman");
        fallbackNoteRun.setColor("666666");
        fallbackNoteRun.setItalic(true);

        int endpointNum = 1;
        for (PathInfo pathInfo : diagnostics.pathInfos) {
            OpenApiSpec.Path pathItem = allPaths.get(pathInfo.path);
            if (pathItem == null) continue;

            Map<String, OpenApiSpec.Operation> operations = extractOperationsFromPath(pathItem);
            if (operations != null && !operations.isEmpty()) {
                for (Map.Entry<String, OpenApiSpec.Operation> opEntry : operations.entrySet()) {
                    OpenApiSpec.Operation operation = opEntry.getValue();
                    if (operation != null) {
                        EndpointOperation endpointOp = new EndpointOperation();
                        endpointOp.path = pathInfo.path;
                        endpointOp.method = opEntry.getKey() != null ? opEntry.getKey().toUpperCase() : "UNKNOWN";
                        endpointOp.operation = operation;

                        addEndpointDetails(document, endpointNum, endpointOp);
                        endpointNum++;
                    }
                }
            }
        }
    }

    private void addGroupSection(XWPFDocument document, int groupNum, String groupName,
                                 List<EndpointOperation> operations) {
        XWPFParagraph groupTitle = document.createParagraph();
        groupTitle.setStyle("Heading2");
        groupTitle.setSpacingBefore(300);
        groupTitle.setSpacingAfter(100);

        XWPFRun groupTitleRun = groupTitle.createRun();
        groupTitleRun.setText(groupNum + ". " + groupName + " (" + operations.size() + " операций)");
        groupTitleRun.setBold(true);
        groupTitleRun.setFontSize(14);
        groupTitleRun.setFontFamily("Times New Roman");
        groupTitleRun.setColor("000000");

        int endpointNum = 1;
        for (EndpointOperation endpointOp : operations) {
            addEndpointDetails(document, endpointNum, endpointOp);
            endpointNum++;
        }
    }

    private void addEndpointDetails(XWPFDocument document, int endpointNum, EndpointOperation endpointOp) {
        XWPFParagraph endpointTitle = document.createParagraph();
        endpointTitle.setSpacingBefore(150);
        endpointTitle.setSpacingAfter(50);
        endpointTitle.setIndentationLeft(360);

        XWPFRun numRun = endpointTitle.createRun();
        numRun.setText(endpointNum + ". ");
        numRun.setBold(true);
        numRun.setFontSize(11);
        numRun.setFontFamily("Times New Roman");

        XWPFRun methodRun = endpointTitle.createRun();
        String method = endpointOp.method != null ? endpointOp.method : "UNKNOWN";
        methodRun.setText(method + " ");
        methodRun.setBold(true);
        methodRun.setFontSize(11);
        methodRun.setFontFamily("Courier New");
        methodRun.setColor(getMethodColor(method));

        XWPFRun pathRun = endpointTitle.createRun();
        pathRun.setText(endpointOp.path != null ? endpointOp.path : "UNKNOWN");
        pathRun.setFontSize(11);
        pathRun.setFontFamily("Courier New");
        pathRun.setColor("000000");
        pathRun.addBreak();

        OpenApiSpec.Operation operation = endpointOp.operation;
        if (operation == null) {
            XWPFRun nullOpRun = endpointTitle.createRun();
            nullOpRun.setText("   ОШИБКА: Операция не найдена");
            nullOpRun.setFontSize(9);
            nullOpRun.setFontFamily("Times New Roman");
            nullOpRun.setColor("FF0000");
            nullOpRun.addBreak();
            return;
        }

        if (operation.getOperationId() != null) {
            XWPFRun opIdRun = endpointTitle.createRun();
            opIdRun.setText("   ID: " + operation.getOperationId());
            opIdRun.setFontSize(9);
            opIdRun.setFontFamily("Times New Roman");
            opIdRun.setColor("666666");
            opIdRun.addBreak();
        }

        if (operation.getSummary() != null) {
            XWPFParagraph summaryParagraph = document.createParagraph();
            summaryParagraph.setSpacingAfter(30);
            summaryParagraph.setIndentationLeft(720);

            XWPFRun summaryRun = summaryParagraph.createRun();
            summaryRun.setText("   " + operation.getSummary());
            summaryRun.setFontSize(10);
            summaryRun.setFontFamily("Times New Roman");
            summaryRun.setColor("000000");
            summaryRun.setItalic(true);
        }

        if (operation.getDescription() != null) {
            XWPFParagraph descParagraph = document.createParagraph();
            descParagraph.setSpacingAfter(50);
            descParagraph.setIndentationLeft(720);

            XWPFRun descRun = descParagraph.createRun();
            descRun.setText("   " + operation.getDescription());
            descRun.setFontSize(10);
            descRun.setFontFamily("Times New Roman");
            descRun.setColor("000000");
        }

        if (operation.getParameters() != null && !operation.getParameters().isEmpty()) {
            addParametersSection(document, operation.getParameters());
        }

        if (operation.getRequestBody() != null) {
            addRequestBodySection(document, operation.getRequestBody());
        }

        if (operation.getResponses() != null && !operation.getResponses().isEmpty()) {
            addResponsesSection(document, operation.getResponses());
        }

        XWPFParagraph separator = document.createParagraph();
        separator.setSpacingBefore(20);
        separator.setSpacingAfter(100);
        XWPFRun separatorRun = separator.createRun();
        separatorRun.setText("─".repeat(60));
        separatorRun.setFontSize(8);
        separatorRun.setFontFamily("Courier New");
        separatorRun.setColor("CCCCCC");
    }

    private void addEmptyEndpointsMessage(XWPFDocument document) {
        XWPFParagraph emptyMsg = document.createParagraph();
        emptyMsg.setSpacingBefore(200);
        emptyMsg.setIndentationLeft(720);

        XWPFRun emptyRun = emptyMsg.createRun();
        emptyRun.setText("В спецификации API не определены endpoints (paths).");
        emptyRun.setFontSize(11);
        emptyRun.setFontFamily("Times New Roman");
        emptyRun.setColor("666666");
        emptyRun.setItalic(true);
    }

    private String getMethodColor(String method) {
        if (method == null) return "000000";
        return switch (method.toUpperCase()) {
            case "GET" -> "008000";
            case "POST" -> "0000FF";
            case "PUT" -> "FF8C00";
            case "DELETE" -> "FF0000";
            case "PATCH" -> "800080";
            default -> "000000";
        };
    }

    private void addParametersSection(XWPFDocument document, List<OpenApiSpec.Parameter> parameters) {
        XWPFParagraph paramsTitle = document.createParagraph();
        paramsTitle.setSpacingBefore(50);
        paramsTitle.setSpacingAfter(20);
        paramsTitle.setIndentationLeft(720);

        XWPFRun paramsTitleRun = paramsTitle.createRun();
        paramsTitleRun.setText("   Параметры:");
        paramsTitleRun.setBold(true);
        paramsTitleRun.setFontSize(10);
        paramsTitleRun.setFontFamily("Times New Roman");
        paramsTitleRun.setColor("000000");

        for (OpenApiSpec.Parameter param : parameters) {
            if (param == null || param.getName() == null) continue;

            XWPFParagraph paramParagraph = document.createParagraph();
            paramParagraph.setSpacingAfter(20);
            paramParagraph.setIndentationLeft(1080);

            XWPFRun paramRun = paramParagraph.createRun();
            paramRun.setFontFamily("Times New Roman");
            paramRun.setFontSize(9);
            paramRun.setColor("000000");

            paramRun.setText("• ");
            paramRun.setBold(true);
            paramRun.setText(param.getName());
            paramRun.setText(" (");
            paramRun.setText(getLocationText(param.getIn()));
            paramRun.setText(") ");
            if (Boolean.TRUE.equals(param.getRequired())) {
                paramRun.setText("* ");
            }
            paramRun.setBold(false);

            if (param.getSchema() != null) {
                paramRun.setText(" - " + getSchemaType(param.getSchema()));
                paramRun.addBreak();
                paramRun.setText("      ");
            }

            if (param.getDescription() != null && !param.getDescription().trim().isEmpty()) {
                paramRun.setText("Описание: " + param.getDescription());
            }
        }
    }

    private void addRequestBodySection(XWPFDocument document, OpenApiSpec.RequestBody requestBody) {
        if (requestBody == null) return;

        XWPFParagraph bodyTitle = document.createParagraph();
        bodyTitle.setSpacingBefore(50);
        bodyTitle.setSpacingAfter(20);
        bodyTitle.setIndentationLeft(720);

        XWPFRun bodyTitleRun = bodyTitle.createRun();
        bodyTitleRun.setText("   Тело запроса:");
        bodyTitleRun.setBold(true);
        bodyTitleRun.setFontSize(10);
        bodyTitleRun.setFontFamily("Times New Roman");
        bodyTitleRun.setColor("000000");

        XWPFParagraph bodyParagraph = document.createParagraph();
        bodyParagraph.setSpacingAfter(50);
        bodyParagraph.setIndentationLeft(1080);

        XWPFRun bodyRun = bodyParagraph.createRun();
        bodyRun.setFontFamily("Times New Roman");
        bodyRun.setFontSize(9);
        bodyRun.setColor("000000");

        if (requestBody.getDescription() != null && !requestBody.getDescription().trim().isEmpty()) {
            bodyRun.setText("• Описание: " + requestBody.getDescription());
            bodyRun.addBreak();
        }

        if (Boolean.TRUE.equals(requestBody.getRequired())) {
            bodyRun.setText("• Обязательное: Да");
            bodyRun.addBreak();
        }

        if (requestBody.getContent() != null && !requestBody.getContent().isEmpty()) {
            bodyRun.setText("• Типы контента: " +
                    requestBody.getContent().keySet().stream()
                            .filter(Objects::nonNull)
                            .map(Object::toString)
                            .collect(Collectors.joining(", ")));
            bodyRun.addBreak();

            for (Map.Entry<String, OpenApiSpec.MediaType> contentEntry : requestBody.getContent().entrySet()) {
                if (contentEntry.getValue() != null && contentEntry.getValue().getSchema() != null) {
                    bodyRun.setText("• Схема: " + getSchemaType(contentEntry.getValue().getSchema()));
                    bodyRun.addBreak();
                }
            }
        }
    }

    private void addResponsesSection(XWPFDocument document, Map<String, OpenApiSpec.Response> responses) {
        if (responses == null || responses.isEmpty()) return;

        XWPFParagraph responsesTitle = document.createParagraph();
        responsesTitle.setSpacingBefore(50);
        responsesTitle.setSpacingAfter(20);
        responsesTitle.setIndentationLeft(720);

        XWPFRun responsesTitleRun = responsesTitle.createRun();
        responsesTitleRun.setText("   Ответы:");
        responsesTitleRun.setBold(true);
        responsesTitleRun.setFontSize(10);
        responsesTitleRun.setFontFamily("Times New Roman");
        responsesTitleRun.setColor("000000");

        List<Map.Entry<String, OpenApiSpec.Response>> sortedResponses = new ArrayList<>(responses.entrySet());
        sortedResponses.sort(Comparator.comparingInt(e -> {
            try {
                return Integer.parseInt(e.getKey());
            } catch (NumberFormatException ex) {
                return Integer.MAX_VALUE;
            }
        }));

        for (Map.Entry<String, OpenApiSpec.Response> entry : sortedResponses) {
            if (entry.getValue() == null) continue;

            XWPFParagraph responseParagraph = document.createParagraph();
            responseParagraph.setSpacingAfter(20);
            responseParagraph.setIndentationLeft(1080);

            XWPFRun responseRun = responseParagraph.createRun();
            responseRun.setFontFamily("Times New Roman");
            responseRun.setFontSize(9);
            responseRun.setColor("000000");

            responseRun.setText("• HTTP " + entry.getKey() + ": ");
            responseRun.setBold(true);
            String description = entry.getValue().getDescription();
            responseRun.setText(description != null && !description.trim().isEmpty() ? description : "Успешный ответ");
            responseRun.setBold(false);

            OpenApiSpec.Response response = entry.getValue();
            if (response.getContent() != null && !response.getContent().isEmpty()) {
                responseRun.addBreak();
                Set<String> contentTypes = response.getContent().keySet().stream()
                        .filter(Objects::nonNull)
                        .collect(Collectors.toSet());
                responseRun.setText("    Типы контента: " +
                        (contentTypes.isEmpty() ? "не указаны" : String.join(", ", contentTypes)));

                for (Map.Entry<String, OpenApiSpec.MediaType> contentEntry : response.getContent().entrySet()) {
                    if (contentEntry.getValue() != null && contentEntry.getValue().getSchema() != null) {
                        OpenApiSpec.Schema schema = contentEntry.getValue().getSchema();
                        responseRun.addBreak();
                        responseRun.setText("    Схема: " + getSchemaType(schema));
                    }
                }
            }
        }
    }

    private String getLocationText(String location) {
        if (location == null) return "Неизвестно";
        switch (location) {
            case "path": return "Путь";
            case "query": return "Запрос";
            case "header": return "Заголовок";
            case "cookie": return "Куки";
            default: return location;
        }
    }

    private String getSchemaType(OpenApiSpec.Schema schema) {
        if (schema == null) return "не определен";

        if (schema.getRef() != null && !schema.getRef().trim().isEmpty()) {
            return extractSchemaName(schema.getRef()) + " {...}";
        }

        StringBuilder typeInfo = new StringBuilder();
        if (schema.getType() != null && !schema.getType().trim().isEmpty()) {
            typeInfo.append(schema.getType());
        }
        if (schema.getFormat() != null && !schema.getFormat().trim().isEmpty()) {
            typeInfo.append(" ($").append(schema.getFormat()).append(")");
        }
        if (schema.getItems() != null) {
            typeInfo.append("[").append(getSchemaType(schema.getItems())).append("]");
        }

        return !typeInfo.isEmpty() ? typeInfo.toString() : "object";
    }

    private String extractSchemaName(String ref) {
        if (ref == null || ref.trim().isEmpty()) return "";
        Pattern pattern = Pattern.compile(".*/([^/]+)$");
        java.util.regex.Matcher matcher = pattern.matcher(ref.trim());
        return matcher.find() ? matcher.group(1) : ref.trim();
    }

    // === КЛАССЫ ДЛЯ ДИАГНОСТИКИ ===
    private static class DiagnosticInfo {
        int totalPaths = 0;
        int totalTaggedOperations = 0;
        int totalUntaggedOperations = 0;
        List<String> specTags = new ArrayList<>();
        List<PathInfo> pathInfos = new ArrayList<>();
    }

    private static class PathInfo {
        String path;
        int operations = 0;
        List<OperationInfo> operationsInfo = new ArrayList<>();
    }

    private static class OperationInfo {
        String method;
        String operationId;
        String summary;
        List<String> tags = new ArrayList<>();
    }

    private static class EndpointOperation {
        String path;
        String method;
        OpenApiSpec.Operation operation;
    }

    // === СХЕМЫ (остаются без изменений) ===
    private void addSchemasSection(XWPFDocument document, OpenApiSpec.Components components) {
        if (components != null && components.getSchemas() != null && !components.getSchemas().isEmpty()) {
            XWPFParagraph schemasTitle = document.createParagraph();
            schemasTitle.setStyle("Heading1");
            schemasTitle.setSpacingBefore(600);
            schemasTitle.setSpacingAfter(300);

            XWPFRun schemasTitleRun = schemasTitle.createRun();
            schemasTitleRun.setText("5. МОДЕЛИ ДАННЫХ");
            schemasTitleRun.setBold(true);
            schemasTitleRun.setFontSize(16);
            schemasTitleRun.setFontFamily("Times New Roman");
            schemasTitleRun.setColor("000000");

            int schemaNum = 1;
            List<Map.Entry<String, OpenApiSpec.Schema>> sortedSchemas =
                    new ArrayList<>(components.getSchemas().entrySet());
            sortedSchemas.sort((e1, e2) -> e1.getKey().compareToIgnoreCase(e2.getKey()));

            for (Map.Entry<String, OpenApiSpec.Schema> entry : sortedSchemas) {
                addSchemaSection(document, schemaNum, entry.getKey(), entry.getValue());
                schemaNum++;
            }
        }
    }

    private void addSchemaSection(XWPFDocument document, int schemaNum, String name, OpenApiSpec.Schema schema) {
        XWPFParagraph schemaNameParagraph = document.createParagraph();
        schemaNameParagraph.setSpacingBefore(300);
        schemaNameParagraph.setSpacingAfter(100);

        XWPFRun schemaNameRun = schemaNameParagraph.createRun();
        schemaNameRun.setText(schemaNum + ". " + name);
        schemaNameRun.setBold(true);
        schemaNameRun.setFontSize(14);
        schemaNameRun.setFontFamily("Times New Roman");
        schemaNameRun.setColor("000000");

        if (schema.getDescription() != null && !schema.getDescription().trim().isEmpty()) {
            XWPFParagraph schemaDesc = document.createParagraph();
            schemaDesc.setSpacingAfter(100);
            schemaDesc.setIndentationLeft(720);

            XWPFRun schemaDescRun = schemaDesc.createRun();
            schemaDescRun.setText(schema.getDescription());
            schemaDescRun.setItalic(true);
            schemaDescRun.setFontFamily("Times New Roman");
            schemaDescRun.setFontSize(11);
            schemaDescRun.setColor("000000");
        }

        StringBuilder schemaInfo = new StringBuilder();
        if (schema.getType() != null && !schema.getType().trim().isEmpty()) {
            schemaInfo.append("Тип: ").append(schema.getType());
        }
        if (schema.getFormat() != null && !schema.getFormat().trim().isEmpty()) {
            if (!schemaInfo.isEmpty()) schemaInfo.append(" | ");
            schemaInfo.append("Формат: ").append(schema.getFormat());
        }

        if (!schemaInfo.isEmpty()) {
            XWPFParagraph typeInfo = document.createParagraph();
            typeInfo.setSpacingAfter(150);
            typeInfo.setIndentationLeft(720);

            XWPFRun typeInfoRun = typeInfo.createRun();
            typeInfoRun.setText(schemaInfo.toString());
            typeInfoRun.setFontFamily("Times New Roman");
            typeInfoRun.setFontSize(11);
            typeInfoRun.setColor("000000");
        }

        if (schema.getProperties() != null && !schema.getProperties().isEmpty()) {
            createSchemaPropertiesTable(document, schema);
        }

        XWPFParagraph spacer = document.createParagraph();
        spacer.setSpacingAfter(300);
    }

    private void createSchemaPropertiesTable(XWPFDocument document, OpenApiSpec.Schema schema) {
        XWPFTable table = document.createTable(1, 4);
        setupSchemaTableProperties(table);

        XWPFTableRow headerRow = table.getRow(0);
        headerRow.getCell(0).setText("Поле");
        headerRow.getCell(1).setText("Тип");
        headerRow.getCell(2).setText("Обязательное");
        headerRow.getCell(3).setText("Описание");
        styleSchemaTableHeader(headerRow);

        Set<String> requiredFields = schema.getRequired() != null ?
                new HashSet<>(schema.getRequired()) : new HashSet<>();

        List<Map.Entry<String, OpenApiSpec.Schema>> sortedProperties =
                new ArrayList<>(schema.getProperties().entrySet());
        sortedProperties.sort((e1, e2) -> {
            String name1 = e1.getKey() != null ? e1.getKey() : "";
            String name2 = e2.getKey() != null ? e2.getKey() : "";
            return name1.compareToIgnoreCase(name2);
        });

        for (Map.Entry<String, OpenApiSpec.Schema> propertyEntry : sortedProperties) {
            String fieldName = propertyEntry.getKey();
            if (fieldName == null || fieldName.trim().isEmpty()) continue;

            OpenApiSpec.Schema fieldSchema = propertyEntry.getValue();
            String fieldType = getSchemaType(fieldSchema);
            String required = requiredFields.contains(fieldName) ? "Да" : "Нет";
            String description = fieldSchema != null && fieldSchema.getDescription() != null ?
                    fieldSchema.getDescription() : "";

            XWPFTableRow row = table.createRow();
            row.getCell(0).setText(fieldName.trim());
            row.getCell(1).setText(fieldType);
            row.getCell(2).setText(required);
            row.getCell(3).setText(description);

            for (int i = 0; i < 4; i++) {
                styleTableCell(row.getCell(i), false);
            }
        }
    }

    private void setupTableProperties(XWPFTable table) {
        CTTblPr tblPr = table.getCTTbl().addNewTblPr();
        CTTblWidth tblW = tblPr.addNewTblW();
        tblW.setW(BigInteger.valueOf(9000));
        tblW.setType(STTblWidth.PCT);
        table.setTableAlignment(TableRowAlign.LEFT);
    }

    private void setupSchemaTableProperties(XWPFTable table) {
        CTTblPr tblPr = table.getCTTbl().addNewTblPr();
        CTTblWidth tblW = tblPr.addNewTblW();
        tblW.setW(BigInteger.valueOf(9500));
        tblW.setType(STTblWidth.PCT);

        CTTblGrid tblGrid = table.getCTTbl().addNewTblGrid();
        tblGrid.addNewGridCol().setW(BigInteger.valueOf(1500));
        tblGrid.addNewGridCol().setW(BigInteger.valueOf(2500));
        tblGrid.addNewGridCol().setW(BigInteger.valueOf(1500));
        tblGrid.addNewGridCol().setW(BigInteger.valueOf(4000));

        table.setTableAlignment(TableRowAlign.LEFT);
    }

    private void styleTableHeader(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            cell.setColor("E6E6FA");
            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                for (XWPFRun run : paragraph.getRuns()) {
                    run.setBold(true);
                    run.setFontFamily("Times New Roman");
                    run.setFontSize(10);
                    run.setColor("000000");
                }
            }
        }
    }

    private void styleSchemaTableHeader(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            cell.setColor("F0F8FF");
            for (XWPFParagraph paragraph : cell.getParagraphs()) {
                paragraph.setAlignment(ParagraphAlignment.CENTER);
                for (XWPFRun run : paragraph.getRuns()) {
                    run.setBold(true);
                    run.setFontFamily("Times New Roman");
                    run.setFontSize(10);
                    run.setColor("000000");
                }
            }
        }
    }

    private void styleTableCell(XWPFTableCell cell, boolean isHeader) {
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            for (XWPFRun run : paragraph.getRuns()) {
                run.setFontFamily("Times New Roman");
                run.setFontSize(9);
                if (isHeader) {
                    run.setBold(true);
                }
                run.setColor("000000");
            }
        }
    }

    private void addPageBreak(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.addBreak(BreakType.PAGE);
    }

    private void addSectionSpacing(XWPFDocument document) {
        addSectionSpacing(document, 400);
    }

    private void addSectionSpacing(XWPFDocument document, int spacing) {
        XWPFParagraph spacer = document.createParagraph();
        spacer.setSpacingAfter(spacing);
    }

    private String generateFileName(String apiTitle) {
        if (apiTitle == null || apiTitle.trim().isEmpty()) {
            apiTitle = "API";
        }

        String safeTitle = apiTitle.replaceAll("[^a-zA-Z0-9а-яА-Я\\s-]", "_")
                .replaceAll("\\s+", "_")
                .substring(0, Math.min(50, apiTitle.length()));

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        return outputDirectory + "/" + safeTitle + "_API_Documentation_" + timestamp + ".docx";
    }
}