package com.bio.xlreport.console;

import com.bio.xlreport.core.model.CompatibilityMode;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import lombok.Builder;
import lombok.Value;

@Value
@Builder
class ConsoleArgs {
    Path reportXml;
    Path dataJson;
    Path templateFile;
    Path outputFile;
    Path dbProfilesFile;
    String dbUrl;
    String dbUser;
    String dbPassword;
    String dbDriver;
    int dbFetchSize;
    CompatibilityMode compatibilityMode;
    boolean stopOnFinish;
    String userUid;
    String userOrgId;
    String userRoles;
    long perfSlaMs;
    Path dbaPackOut;
    Map<String, String> reportParams;

    static ConsoleArgs parse(String[] args) {
        var b = ConsoleArgs.builder()
            .compatibilityMode(CompatibilityMode.STRICT)
            .stopOnFinish(true)
            .dbDriver("oracle.jdbc.OracleDriver")
            .dbFetchSize(1_000)
            .userUid("localuser")
            .userOrgId("")
            .userRoles("*")
            .perfSlaMs(0L)
            .reportParams(new LinkedHashMap<>());

        for (String arg : args) {
            if (arg == null || arg.isBlank()) {
                continue;
            }
            if (arg.startsWith("/rpt:")) {
                b.reportXml(Path.of(arg.substring("/rpt:".length())));
            } else if (arg.startsWith("--rpt=")) {
                b.reportXml(Path.of(arg.substring("--rpt=".length())));
            } else if (arg.startsWith("/data:")) {
                b.dataJson(Path.of(arg.substring("/data:".length())));
            } else if (arg.startsWith("--data=")) {
                b.dataJson(Path.of(arg.substring("--data=".length())));
            } else if (arg.startsWith("/template:")) {
                b.templateFile(Path.of(arg.substring("/template:".length())));
            } else if (arg.startsWith("--template=")) {
                b.templateFile(Path.of(arg.substring("--template=".length())));
            } else if (arg.startsWith("/out:")) {
                b.outputFile(Path.of(arg.substring("/out:".length())));
            } else if (arg.startsWith("--out=")) {
                b.outputFile(Path.of(arg.substring("--out=".length())));
            } else if (arg.startsWith("/dbProfiles:")) {
                b.dbProfilesFile(Path.of(arg.substring("/dbProfiles:".length())));
            } else if (arg.startsWith("--dbProfiles=")) {
                b.dbProfilesFile(Path.of(arg.substring("--dbProfiles=".length())));
            } else if (arg.startsWith("/dbUrl:")) {
                b.dbUrl(arg.substring("/dbUrl:".length()));
            } else if (arg.startsWith("--dbUrl=")) {
                b.dbUrl(arg.substring("--dbUrl=".length()));
            } else if (arg.startsWith("/dbUser:")) {
                b.dbUser(arg.substring("/dbUser:".length()));
            } else if (arg.startsWith("--dbUser=")) {
                b.dbUser(arg.substring("--dbUser=".length()));
            } else if (arg.startsWith("/dbPassword:")) {
                b.dbPassword(arg.substring("/dbPassword:".length()));
            } else if (arg.startsWith("--dbPassword=")) {
                b.dbPassword(arg.substring("--dbPassword=".length()));
            } else if (arg.startsWith("/dbDriver:")) {
                b.dbDriver(arg.substring("/dbDriver:".length()));
            } else if (arg.startsWith("--dbDriver=")) {
                b.dbDriver(arg.substring("--dbDriver=".length()));
            } else if (arg.startsWith("/dbFetchSize:")) {
                b.dbFetchSize(parsePositiveInt(arg.substring("/dbFetchSize:".length()), 1_000));
            } else if (arg.startsWith("--dbFetchSize=")) {
                b.dbFetchSize(parsePositiveInt(arg.substring("--dbFetchSize=".length()), 1_000));
            } else if (arg.startsWith("/mode:")) {
                b.compatibilityMode(parseMode(arg.substring("/mode:".length())));
            } else if (arg.startsWith("--mode=")) {
                b.compatibilityMode(parseMode(arg.substring("--mode=".length())));
            } else if (arg.startsWith("/rptPrms:")) {
                b.reportParams(parseParams(arg.substring("/rptPrms:".length())));
            } else if (arg.startsWith("/rptUserUID:")) {
                b.userUid(arg.substring("/rptUserUID:".length()));
            } else if (arg.startsWith("/rptUserOrgId:")) {
                b.userOrgId(arg.substring("/rptUserOrgId:".length()));
            } else if (arg.startsWith("/rptUserRoles:")) {
                b.userRoles(arg.substring("/rptUserRoles:".length()));
            } else if (arg.startsWith("/rptStopOnFinish:")) {
                b.stopOnFinish(parseBoolean(arg.substring("/rptStopOnFinish:".length()), true));
            } else if (arg.startsWith("/perfSlaMs:")) {
                b.perfSlaMs(parsePositiveLong(arg.substring("/perfSlaMs:".length()), 0L));
            } else if (arg.startsWith("--perfSlaMs=")) {
                b.perfSlaMs(parsePositiveLong(arg.substring("--perfSlaMs=".length()), 0L));
            } else if (arg.startsWith("/dbaPackOut:")) {
                b.dbaPackOut(Path.of(arg.substring("/dbaPackOut:".length())));
            } else if (arg.startsWith("--dbaPackOut=")) {
                b.dbaPackOut(Path.of(arg.substring("--dbaPackOut=".length())));
            }
        }
        return b.build();
    }

    private static CompatibilityMode parseMode(String raw) {
        if (raw == null) {
            return CompatibilityMode.STRICT;
        }
        return "lenient".equalsIgnoreCase(raw.trim()) ? CompatibilityMode.LENIENT : CompatibilityMode.STRICT;
    }

    private static Map<String, String> parseParams(String raw) {
        Map<String, String> params = new LinkedHashMap<>();
        if (raw == null || raw.isBlank()) {
            return params;
        }
        String[] parts = raw.split(";");
        for (String p : parts) {
            int eq = p.indexOf('=');
            if (eq <= 0 || eq >= p.length() - 1) {
                continue;
            }
            params.put(p.substring(0, eq).trim(), p.substring(eq + 1).trim());
        }
        return params;
    }

    private static boolean parseBoolean(String raw, boolean def) {
        if (raw == null || raw.isBlank()) {
            return def;
        }
        String v = raw.trim().toLowerCase(Locale.ROOT);
        return v.equals("true") || v.equals("yes") || v.equals("1") || v.equals("t") || v.equals("y");
    }

    private static int parsePositiveInt(String raw, int def) {
        if (raw == null || raw.isBlank()) {
            return def;
        }
        try {
            int parsed = Integer.parseInt(raw.trim());
            return parsed > 0 ? parsed : def;
        } catch (NumberFormatException ex) {
            return def;
        }
    }

    private static long parsePositiveLong(String raw, long def) {
        if (raw == null || raw.isBlank()) {
            return def;
        }
        try {
            long parsed = Long.parseLong(raw.trim());
            return parsed > 0 ? parsed : def;
        } catch (NumberFormatException ex) {
            return def;
        }
    }
}
