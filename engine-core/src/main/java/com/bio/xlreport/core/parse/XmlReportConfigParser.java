package com.bio.xlreport.core.parse;

import com.bio.xlreport.core.model.CompatibilityMode;
import com.bio.xlreport.core.model.DataSourceConfig;
import com.bio.xlreport.core.model.ExtAttributes;
import com.bio.xlreport.core.model.FieldDef;
import com.bio.xlreport.core.model.PostScriptConfig;
import com.bio.xlreport.core.model.ReportConfig;
import com.bio.xlreport.core.model.SortDef;
import com.bio.xlreport.core.model.SortDirection;
import java.io.StringReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;
import javax.xml.parsers.DocumentBuilderFactory;
import lombok.extern.slf4j.Slf4j;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

@Slf4j
public class XmlReportConfigParser {

    public ReportConfig parse(String xml, CompatibilityMode mode) throws Exception {
        return parse(xml, null, mode);
    }

    public ReportConfig parse(String xml, Path sourceFile, CompatibilityMode mode) throws Exception {
        var dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(false);
        dbf.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        dbf.setFeature("http://xml.org/sax/features/external-general-entities", false);
        dbf.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        Document doc = dbf.newDocumentBuilder().parse(new InputSource(new StringReader(xml)));
        Element root = doc.getDocumentElement();
        boolean reportXsdStyle = "report".equalsIgnoreCase(root.getTagName());

        Path templatePath = resolveTemplatePath(root, sourceFile);
        Path outputPath = resolveOutputPath(root, sourceFile);
        String fullCode = attr(root, "full_code", deriveFullCode(sourceFile));
        if (fullCode == null || fullCode.isBlank()) {
            fullCode = "legacy.report";
        }

        var builder = ReportConfig.builder()
            .uid(UUID.randomUUID().toString())
            .fullCode(fullCode)
            .title(text(root, "title", ""))
            .subject(text(root, "subject", ""))
            .author(text(root, "autor", ""))
            .templatePath(templatePath)
            .outputPath(outputPath)
            .compatibilityMode(mode)
            .convertResultToPdf(boolAttr(root, "convertResultToPDF", false))
            .extAttributes(parseExtAttributes(root));

        validateLegacyFlags(root, mode);

        parseParams(root, "params", builder);
        parseParams(root, "inParams", builder);
        parseEnv(root, reportXsdStyle, builder);
        parseDataSources(root, builder, mode);
        parsePostScripts(root, builder, mode);

        return builder.build();
    }

    public ReportConfig parseFromFile(Path sourceFile, CompatibilityMode mode) throws Exception {
        String xml = Files.readString(sourceFile);
        return parse(xml, sourceFile, mode);
    }

    private void parseDataSources(Element root, ReportConfig.ReportConfigBuilder builder, CompatibilityMode mode) {
        NodeList dsNodes = root.getElementsByTagName("ds");
        for (int i = 0; i < dsNodes.getLength(); i++) {
            Element ds = (Element) dsNodes.item(i);
            var dsBuilder = DataSourceConfig.builder()
                .rangeName(attr(ds, "range", ""))
                .title(attr(ds, "title", attr(ds, "range", "")))
                .connectionName(attr(ds, "connectionName", "default"))
                .singleRow(boolAttr(ds, "singleRow", false))
                .leaveGroupData(boolAttr(ds, "leaveGroupData", false))
                .timeoutMinutes(intAttr(ds, "timeoutMinutes", 90))
                .maxRowsLimit(longAttr(ds, "maxRowsLimit", 900_000L));

            Element sqlElem = firstChild(ds, "sql");
            if (sqlElem != null) {
                dsBuilder.sql(sqlElem.getTextContent());
                String commandType = attr(sqlElem, "commandType", "Text");
                dsBuilder.commandType(commandType);
                if (!"text".equalsIgnoreCase(commandType)) {
                    String msg = "Unsupported ds/sql commandType='" + commandType + "' for range '" + attr(ds, "range", "") + "'.";
                    if (mode == CompatibilityMode.STRICT) {
                        throw new IllegalArgumentException(msg + " Use LENIENT or migrate commandType to Text.");
                    }
                    log.warn("{} Proceeding in LENIENT mode.", msg);
                }
            }

            Element fieldsElem = firstChild(ds, "fields");
            if (fieldsElem != null) {
                NodeList fieldNodes = fieldsElem.getElementsByTagName("field");
                for (int k = 0; k < fieldNodes.getLength(); k++) {
                    Element f = (Element) fieldNodes.item(k);
                    if (!boolAttr(f, "expEnabled", true) || boolAttr(f, "hidden", false)) {
                        continue;
                    }
                    dsBuilder.field(
                        FieldDef.builder()
                            .name(attr(f, "name", ""))
                            .type(attr(f, "type", "string"))
                            .align(attr(f, "align", "left"))
                            .header(attr(f, "header", ""))
                            .format(attr(f, "expFormat", "@"))
                            .width(intAttr(f, "expWidth", 20))
                            .build()
                    );
                }
            }

            Element sortsElem = firstChild(ds, "sorts");
            if (sortsElem != null) {
                NodeList sortNodes = sortsElem.getElementsByTagName("sort");
                for (int k = 0; k < sortNodes.getLength(); k++) {
                    Element s = (Element) sortNodes.item(k);
                    String dir = attr(s, "direction", "ASC").toUpperCase();
                    dsBuilder.sort(
                        SortDef.builder()
                            .fieldName(attr(s, "fieldName", ""))
                            .direction("DESC".equals(dir) ? SortDirection.DESC : SortDirection.ASC)
                            .build()
                    );
                }
            }

            builder.dataSource(dsBuilder.build());
        }
    }

    private void validateLegacyFlags(Element root, CompatibilityMode mode) {
        boolean convertToPdf = boolAttr(root, "convertResultToPDF", false);
        Element append = firstChild(root, "append");
        String pwdOpen = text(append, "pwdOpen", "").trim();
        String pwdWrite = text(append, "pwdWrite", "").trim();
        boolean liveScripts = boolAttr(root, "liveScripts", false);

        if (convertToPdf) {
            String msg = "convertResultToPDF is not implemented in Java runtime.";
            if (mode == CompatibilityMode.STRICT) {
                throw new IllegalArgumentException(msg + " Disable flag or use LENIENT.");
            }
            log.warn("{} Running without PDF conversion.", msg);
        }
        if (!pwdOpen.isBlank() || !pwdWrite.isBlank()) {
            String msg = "Workbook passwords (pwdOpen/pwdWrite) are not implemented.";
            if (mode == CompatibilityMode.STRICT) {
                throw new IllegalArgumentException(msg + " Remove passwords from XML or use LENIENT.");
            }
            log.warn("{} Running without password protection.", msg);
        }
        if (liveScripts) {
            log.warn("liveScripts=true detected. Execution is controlled by configured postScripts/legacy hooks only.");
        }
    }

    private void parsePostScripts(
        Element root,
        ReportConfig.ReportConfigBuilder builder,
        CompatibilityMode mode
    ) {
        Element scripts = firstChild(root, "postScripts");
        if (scripts != null) {
            NodeList scriptNodes = scripts.getElementsByTagName("script");
            for (int i = 0; i < scriptNodes.getLength(); i++) {
                Element s = (Element) scriptNodes.item(i);
                builder.postScript(
                    PostScriptConfig.builder()
                        .name(attr(s, "name", "script-" + i))
                        .scriptPath(attr(s, "path", null))
                        .inlineScript(s.getTextContent())
                        .timeoutMs(longAttr(s, "timeoutMs", 30_000L))
                        .build()
                );
            }
        }

        // Backward compatibility for legacy VBA entries.
        String macroBefore = text(root, "macroBefore", "").trim();
        Element macroAfterElem = firstChild(root, "macroAfter");
        String macroAfter = macroAfterElem != null ? attr(macroAfterElem, "name", text(root, "macroAfter", "")).trim() : text(root, "macroAfter", "").trim();
        String autoStart = text(root, "autostart", "").trim();
        if (!macroBefore.isBlank()) {
            builder.legacyMacroBefore(macroBefore);
        }
        if (!macroAfter.isBlank()) {
            builder.legacyMacroAfter(macroAfter);
        }
        if (!autoStart.isBlank()) {
            builder.legacyAutostart(autoStart);
        }
        if (!macroBefore.isBlank() || !macroAfter.isBlank() || !autoStart.isBlank()) {
            log.warn("Legacy VBA fields found. macroBefore='{}', macroAfter='{}', autostart='{}'.", macroBefore, macroAfter, autoStart);
        }

        // Legacy SQL hooks are captured as executable runtime stages.
        Element sqlBefore = firstChild(root, "sqlScriptBefore");
        if (sqlBefore != null && !sqlBefore.getTextContent().isBlank()) {
            builder.preSqlScript(sqlBefore.getTextContent());
        }
        Element sqlAfter = firstChild(root, "sqlScriptAfter");
        if (sqlAfter != null && !sqlAfter.getTextContent().isBlank()) {
            builder.postSqlScript(sqlAfter.getTextContent());
        }
        Element sqlScripts = firstChild(root, "sqlScripts");
        if (sqlScripts != null) {
            NodeList sqlScriptNodes = sqlScripts.getElementsByTagName("script");
            for (int i = 0; i < sqlScriptNodes.getLength(); i++) {
                Element script = (Element) sqlScriptNodes.item(i);
                String body = script.getTextContent();
                if (body == null || body.isBlank()) {
                    continue;
                }
                String stage = attr(script, "stage", attr(script, "hook", attr(script, "type", "after")));
                if ("before".equalsIgnoreCase(stage) || "pre".equalsIgnoreCase(stage)) {
                    builder.preSqlScript(body);
                } else {
                    builder.postSqlScript(body);
                }
            }
        }
    }

    private void parseParams(Element root, String sectionName, ReportConfig.ReportConfigBuilder builder) {
        Element params = firstChild(root, sectionName);
        if (params == null) {
            return;
        }
        NodeList paramNodes = params.getElementsByTagName("param");
        for (int i = 0; i < paramNodes.getLength(); i++) {
            Element p = (Element) paramNodes.item(i);
            builder.param(attr(p, "name", "param" + i), p.getTextContent());
        }
    }

    private void parseEnv(Element root, boolean reportXsdStyle, ReportConfig.ReportConfigBuilder builder) {
        if (reportXsdStyle) {
            Element env = firstChild(root, "environment");
            if (env == null) {
                return;
            }
            NodeList vars = env.getElementsByTagName("variable");
            for (int i = 0; i < vars.getLength(); i++) {
                Element v = (Element) vars.item(i);
                builder.env(attr(v, "name", "env" + i), v.getTextContent());
            }
            return;
        }
        Element append = firstChild(root, "append");
        if (append == null) {
            return;
        }
        Element env = firstChild(append, "envVars");
        if (env == null) {
            return;
        }
        NodeList vars = env.getElementsByTagName("var");
        for (int i = 0; i < vars.getLength(); i++) {
            Element v = (Element) vars.item(i);
            builder.env(attr(v, "name", "env" + i), v.getTextContent());
        }
    }

    private ExtAttributes parseExtAttributes(Element root) {
        Element append = firstChild(root, "append");
        return ExtAttributes.builder()
            .targetFormat(attr(root, "targetFormat", "msexcel"))
            .liveScripts(boolAttr(root, "liveScripts", false))
            .userLogin(text(append, "userName", ""))
            .userUID(text(append, "userUID", ""))
            .remoteIP(text(append, "remoteIP", ""))
            .shortCode(text(append, "shortCode", "report"))
            .workPath(text(append, "rptWorkPath", ""))
            .localPath(text(append, "rptLocalPath", ""))
            .pwdOpen(text(append, "pwdOpen", ""))
            .pwdWrite(text(append, "pwdWrite", ""))
            .build();
    }

    private Path resolveTemplatePath(Element root, Path sourceFile) {
        String direct = text(root, "adv_template", "").trim();
        if (!direct.isBlank()) {
            return Path.of(direct);
        }
        if (sourceFile != null) {
            String fileName = sourceFile.getFileName().toString();
            String base = fileName.replace("(rpt).xml", "").replace(".xml", "");
            Path parent = sourceFile.getParent();
            if (parent != null) {
                return parent.resolve(base + ".xlsx");
            }
        }
        return Path.of("template.xlsx");
    }

    private Path resolveOutputPath(Element root, Path sourceFile) {
        String output = text(root, "filename_fmt", "").trim();
        if (!output.isBlank()) {
            return Path.of(output);
        }
        if (sourceFile != null) {
            String fileName = sourceFile.getFileName().toString().replace("(rpt).xml", "").replace(".xml", "");
            Path parent = sourceFile.getParent();
            if (parent != null) {
                return parent.resolve(fileName + "-result.xlsx");
            }
        }
        return Path.of("report.xlsx");
    }

    private String deriveFullCode(Path sourceFile) {
        if (sourceFile == null) {
            return "legacy.report";
        }
        String f = sourceFile.getFileName().toString();
        return f.replace("(rpt).xml", "").replace(".xml", "");
    }

    private static Element firstChild(Element parent, String name) {
        if (parent == null) {
            return null;
        }
        NodeList list = parent.getElementsByTagName(name);
        if (list.getLength() == 0) {
            return null;
        }
        for (int i = 0; i < list.getLength(); i++) {
            Node n = list.item(i);
            if (n.getParentNode() == parent) {
                return (Element) n;
            }
        }
        return (Element) list.item(0);
    }

    private static String text(Element parent, String name, String def) {
        Element child = firstChild(parent, name);
        return child == null ? def : child.getTextContent();
    }

    private static String attr(Element elem, String name, String def) {
        if (elem == null || !elem.hasAttribute(name)) {
            return def;
        }
        String value = elem.getAttribute(name);
        return value == null || value.isBlank() ? def : value;
    }

    private static boolean boolAttr(Element elem, String name, boolean def) {
        String raw = attr(elem, name, Boolean.toString(def));
        return "true".equalsIgnoreCase(raw) || "1".equals(raw);
    }

    private static int intAttr(Element elem, String name, int def) {
        try {
            return Integer.parseInt(attr(elem, name, Integer.toString(def)));
        } catch (NumberFormatException ex) {
            return def;
        }
    }

    private static long longAttr(Element elem, String name, long def) {
        try {
            return Long.parseLong(attr(elem, name, Long.toString(def)));
        } catch (NumberFormatException ex) {
            return def;
        }
    }
}
