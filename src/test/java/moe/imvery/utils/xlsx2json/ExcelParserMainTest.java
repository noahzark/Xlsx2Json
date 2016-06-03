package moe.imvery.utils.xlsx2json;

import org.junit.Test;
import org.skyscreamer.jsonassert.JSONAssert;

import java.io.BufferedReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.junit.Assert.*;

/**
 * Created by Feliciano on 6/2/2016.
 */
public class ExcelParserMainTest {

    @Test
    public void main() throws Exception {
        ExcelParserMain.main(new String[]{"testcases/test.xlsx", "map", "true"});

        Path path = Paths.get("testcases/test.json");
        BufferedReader reader = Files.newBufferedReader(path);
        String jsonText = reader.readLine();

        path = Paths.get("testcases/test3.expected.json");
        reader = Files.newBufferedReader(path);
        String expected = reader.readLine();

        JSONAssert.assertEquals(expected, jsonText, false);
    }

    @Test
    public void mainMultisheets() throws Exception {
        ExcelParserMain.main(new String[]{"testcases/test.xlsx", "weaponStuffs skillStuffs essenceStuffs", "true"});

        Path path = Paths.get("testcases/test.json");
        BufferedReader reader = Files.newBufferedReader(path);
        String jsonText = reader.readLine();

        path = Paths.get("testcases/test4.expected.json");
        reader = Files.newBufferedReader(path);
        String expected = reader.readLine();

        JSONAssert.assertEquals(expected, jsonText, false);
    }

    @Test
    public void mainHideSheetName() throws Exception {
        ExcelParserMain.main(new String[]{"testcases/test.xlsx", "map", "false"});

        Path path = Paths.get("testcases/test.json");
        BufferedReader reader = Files.newBufferedReader(path);
        String jsonText = reader.readLine();

        path = Paths.get("testcases/test5.expected.json");
        reader = Files.newBufferedReader(path);
        String expected = reader.readLine();

        JSONAssert.assertEquals(expected, jsonText, false);
    }

    @Test
    public void mainMultiSheetsAndHideSheetName() throws Exception {
        ExcelParserMain.main(new String[]{"testcases/test.xlsx", "map monster", "false"});

        Path path = Paths.get("testcases/test.json");
        BufferedReader reader = Files.newBufferedReader(path);
        String jsonText = reader.readLine();

        path = Paths.get("testcases/test6.expected.json");
        reader = Files.newBufferedReader(path);
        String expected = reader.readLine();

        JSONAssert.assertEquals(expected, jsonText, false);
    }

    @Test(expected=IllegalArgumentException.class)
    public void mainWrongTargetName() throws Exception {
        ExcelParserMain.main(new String[]{"test.xls", "", ""});
    }

    @Test(expected=IllegalArgumentException.class)
    public void mainWrongNumberofArguments() throws Exception {
        ExcelParserMain.main(new String[]{"test.xls"});
    }

}