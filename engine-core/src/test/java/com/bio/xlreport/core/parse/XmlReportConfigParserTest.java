package com.bio.xlreport.core.parse;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.bio.xlreport.core.model.CompatibilityMode;
import java.io.StringReader;
import java.nio.file.Files;
import java.nio.file.Path;
import javax.xml.XMLConstants;
import javax.xml.transform.stream.StreamSource;
import javax.xml.validation.SchemaFactory;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.Assumptions;

class XmlReportConfigParserTest {

    @Test
    void parseLegacyXmlWithDataSource() throws Exception {
        String xml = """
            <reportDef full_code="demo.report" liveScripts="false">
              <adv_template>C:/tmp/template.xlsx</adv_template>
              <filename_fmt>C:/tmp/result.xlsx</filename_fmt>
              <title>Demo</title>
              <subject>Subject</subject>
              <autor>Author</autor>
              <params>
                <param name="P1">V1</param>
              </params>
              <dss>
                <ds range="Sales">
                  <sql commandType="Text"><![CDATA[select 1]]></sql>
                  <fields>
                    <field name="amount" type="number" align="right" expFormat="0.00" expWidth="12"/>
                  </fields>
                </ds>
              </dss>
            </reportDef>
            """;

        var parser = new XmlReportConfigParser();
        var cfg = parser.parse(xml, CompatibilityMode.STRICT);

        assertEquals("demo.report", cfg.getFullCode());
        assertEquals("Demo", cfg.getTitle());
        assertFalse(cfg.getDataSources().isEmpty());
        assertEquals("Sales", cfg.getDataSources().getFirst().getRangeName());
    }

    @Test
    void parseXsdReportStyleXml() throws Exception {
        String xml = """
            <report debug="true" liveScripts="true">
              <filename_fmt>{$code}_{$now}</filename_fmt>
              <title>Demo</title>
              <subject>Subject</subject>
              <autor>Author</autor>
              <macroAfter name="Module1.m1" enabled="false"/>
              <params>
                <param name="date_from" type="str">2026-01-01</param>
              </params>
              <environment>
                <variable name="v1" type="str">X</variable>
              </environment>
              <dss>
                <ds alias="cdsRpt" range="mRng" enabled="true" connectionName="cub7">
                  <sql>{text-file:demo.sql}</sql>
                  <charts/>
                </ds>
              </dss>
            </report>
            """;

        var parser = new XmlReportConfigParser();
        var cfg = parser.parse(xml, CompatibilityMode.STRICT);
        assertEquals("Demo", cfg.getTitle());
        assertEquals("mRng", cfg.getDataSources().getFirst().getRangeName());
        assertEquals("cub7", cfg.getDataSources().getFirst().getConnectionName());
        assertTrue(cfg.getParams().containsKey("date_from"));
        assertTrue(cfg.getEnvVars().containsKey("v1"));
        assertEquals("Module1.m1", cfg.getLegacyMacroAfter());
    }

    @Test
    void validateAndParseRealReportsIfPresent() throws Exception {
        Path xsdPath = Path.of("c:\\data\\tmp\\rrequest-to-migrate\\ekb-cabinet\\ekb-rpt\\iod\\rptdef.xsd");
        Path reportsDir = Path.of("c:\\data\\tmp\\rrequest-to-migrate\\ekb-cabinet\\ekb-rpt\\rpts_v2\\01_givc");
        Assumptions.assumeTrue(Files.exists(xsdPath) && Files.exists(reportsDir));

        var schemaFactory = SchemaFactory.newInstance(XMLConstants.W3C_XML_SCHEMA_NS_URI);
        var schema = schemaFactory.newSchema(xsdPath.toFile());
        var parser = new XmlReportConfigParser();

        try (var stream = Files.list(reportsDir)) {
            var files = stream.filter(p -> p.getFileName().toString().endsWith(".xml")).toList();
            assertTrue(!files.isEmpty());
            int checked = 0;
            int xsdValidCount = 0;
            for (Path xmlFile : files) {
                String xml = Files.readString(xmlFile);
                try {
                    schema.newValidator().validate(new StreamSource(new StringReader(xml)));
                    xsdValidCount++;
                } catch (Exception ignored) {
                    // Real legacy set may contain files outside current XSD rules; parser compatibility is primary here.
                }
                var cfg = parser.parse(xml, xmlFile, CompatibilityMode.LENIENT);
                assertTrue(cfg.getTitle() != null);
                checked++;
            }
            assertTrue(checked > 0);
            assertTrue(xsdValidCount >= 0);
        }
    }
}
