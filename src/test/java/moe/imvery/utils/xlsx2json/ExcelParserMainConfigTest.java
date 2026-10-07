package moe.imvery.utils.xlsx2json;

import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;
import org.skyscreamer.jsonassert.JSONAssert;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.Assert.assertTrue;

public class ExcelParserMainConfigTest {

    @Rule
    public TemporaryFolder folder = new TemporaryFolder();

    /**
     * Copy the test workbook into a temp folder, run main() with a config file, and return the generated JSON
     */
    private String runWithConfig(String configJson) throws Exception {
        File xlsx = new File(folder.getRoot(), "test.xlsx");
        Files.copy(Paths.get("testcases/test.xlsx"), xlsx.toPath());

        File config = folder.newFile("config.json");
        String content = configJson.replace("TARGET", xlsx.getAbsolutePath().replace("\\", "\\\\"));
        Files.write(config.toPath(), content.getBytes(StandardCharsets.UTF_8));

        ExcelParserMain.main(new String[]{"-config=" + config.getAbsolutePath()});

        Path output = new File(folder.getRoot(), "test.json").toPath();
        assertTrue("Expected " + output + " to be written", Files.exists(output));
        return new String(Files.readAllBytes(output), StandardCharsets.UTF_8);
    }

    private static String expected(String name) throws Exception {
        return new String(Files.readAllBytes(Paths.get("testcases/" + name)), StandardCharsets.UTF_8);
    }

    @Test
    public void configWithShowTrue() throws Exception {
        String json = runWithConfig("{\"target\": \"TARGET\", \"sheet\": \"map\", \"show\": \"true\"}");
        JSONAssert.assertEquals(expected("test3.expected.json"), json, false);
    }

    @Test
    public void configWithBooleanShow() throws Exception {
        String json = runWithConfig("{\"target\": \"TARGET\", \"sheet\": \"map\", \"show\": true}");
        JSONAssert.assertEquals(expected("test3.expected.json"), json, false);
    }

    @Test
    public void configWithoutShow() throws Exception {
        String json = runWithConfig("{\"target\": \"TARGET\", \"sheet\": \"map\"}");
        JSONAssert.assertEquals(expected("test5.expected.json"), json, false);
    }
}
