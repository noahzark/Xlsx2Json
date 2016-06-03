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

        ExcelParserMain.validateJson("testcases/test3.expected.json", "testcases/test.json", false);
    }

    @Test
    public void mainMultisheets() throws Exception {
        ExcelParserMain.main(new String[]{"testcases/test.xlsx", "weaponStuffs skillStuffs essenceStuffs", "true"});

        ExcelParserMain.validateJson("testcases/test4.expected.json", "testcases/test.json", false);
    }

    @Test
    public void mainHideSheetName() throws Exception {
        ExcelParserMain.main(new String[]{"testcases/test.xlsx", "map", "false"});

        ExcelParserMain.validateJson("testcases/test5.expected.json", "testcases/test.json", false);
    }

    @Test
    public void mainMultiSheetsAndHideSheetName() throws Exception {
        ExcelParserMain.main(new String[]{"testcases/test.xlsx", "map monster", "false"});

        ExcelParserMain.validateJson("testcases/test6.expected.json", "testcases/test.json", false);
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