package com.bio.xlreport.console;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.bio.xlreport.core.model.CompatibilityMode;
import org.junit.jupiter.api.Test;

class ConsoleArgsTest {

    @Test
    void parseLegacyStyleArgs() {
        String[] args = {
            "/rpt:C:/rpts/form100(rpt).xml",
            "/template:C:/rpts/form100-template.xlsx",
            "/data:C:/tmp/data.json",
            "/mode:lenient",
            "/rptPrms:date_from=2026-01-01;date_to=2026-02-01",
            "/rptUserUID:U1",
            "/rptUserOrgId:ORG1",
            "/rptUserRoles:admin",
            "/rptStopOnFinish:false"
        };

        ConsoleArgs parsed = ConsoleArgs.parse(args);
        assertEquals("C:\\rpts\\form100(rpt).xml", parsed.getReportXml().toString());
        assertEquals("C:\\rpts\\form100-template.xlsx", parsed.getTemplateFile().toString());
        assertEquals("C:\\tmp\\data.json", parsed.getDataJson().toString());
        assertEquals(CompatibilityMode.LENIENT, parsed.getCompatibilityMode());
        assertEquals("2026-01-01", parsed.getReportParams().get("date_from"));
        assertEquals("U1", parsed.getUserUid());
        assertEquals("ORG1", parsed.getUserOrgId());
        assertEquals("admin", parsed.getUserRoles());
        assertFalse(parsed.isStopOnFinish());
    }

    @Test
    void parseBooleanVariants() {
        ConsoleArgs parsed = ConsoleArgs.parse(new String[] { "/rpt:a.xml", "/rptStopOnFinish:yes" });
        assertTrue(parsed.isStopOnFinish());
    }

    @Test
    void parseOracleArgs() {
        String[] args = {
            "/rpt:C:/rpts/form100(rpt).xml",
            "/dbUrl:jdbc:oracle:thin:@//10.10.10.10:1521/XEPDB1",
            "/dbUser:RPT_USER",
            "/dbPassword:secret",
            "/dbFetchSize:2000"
        };

        ConsoleArgs parsed = ConsoleArgs.parse(args);
        assertEquals("jdbc:oracle:thin:@//10.10.10.10:1521/XEPDB1", parsed.getDbUrl());
        assertEquals("RPT_USER", parsed.getDbUser());
        assertEquals("secret", parsed.getDbPassword());
        assertEquals("oracle.jdbc.OracleDriver", parsed.getDbDriver());
        assertEquals(2000, parsed.getDbFetchSize());
    }
}
