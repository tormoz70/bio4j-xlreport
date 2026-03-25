package com.bio.xlreport.poi.golden;

import com.bio.xlreport.poi.golden.spec.GoldenCaseSpec;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Stream;

public final class GoldenCaseRepository {
    private static final ObjectMapper MAPPER = new ObjectMapper()
        .configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);

    private GoldenCaseRepository() {
    }

    public static List<GoldenCaseSpec> loadAllFromProject(Path projectRoot) throws Exception {
        Path dir = null;
        var resource = GoldenCaseRepository.class.getClassLoader().getResource("golden-cases");
        if (resource != null && "file".equalsIgnoreCase(resource.getProtocol())) {
            dir = Path.of(resource.toURI());
        }
        if (dir == null || !Files.exists(dir)) {
            Path candidateA = projectRoot.resolve("src").resolve("test").resolve("resources").resolve("golden-cases");
            Path candidateB = projectRoot.resolve("engine-poi").resolve("src").resolve("test").resolve("resources").resolve("golden-cases");
            if (Files.exists(candidateA)) {
                dir = candidateA;
            } else if (Files.exists(candidateB)) {
                dir = candidateB;
            } else {
                return List.of();
            }
        }
        List<GoldenCaseSpec> specs = new ArrayList<>();
        try (Stream<Path> stream = Files.list(dir)) {
            stream
                .filter(p -> p.getFileName().toString().endsWith(".json"))
                .sorted(Comparator.comparing(p -> p.getFileName().toString()))
                .forEach(path -> {
                    try {
                        specs.add(MAPPER.readValue(path.toFile(), GoldenCaseSpec.class));
                    } catch (Exception ex) {
                        throw new RuntimeException("Failed to parse golden case: " + path, ex);
                    }
                });
        }
        return specs;
    }
}
