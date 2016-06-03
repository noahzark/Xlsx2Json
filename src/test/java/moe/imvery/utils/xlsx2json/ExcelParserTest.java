package moe.imvery.utils.xlsx2json;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Test;
import org.skyscreamer.jsonassert.JSONAssert;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Created by Feliciano on 6/2/2016.
 */
public class ExcelParserTest {

    @Test
    public void parseSheetWithAllBasicTypes() throws Exception {
        String fileName = "testcases/test";
        File excelFile = new File(fileName + ".xlsx");

        try(FileInputStream inp = new FileInputStream( excelFile )) {
            Workbook workbook = WorkbookFactory.create(inp);
            // Start constructing JSON.
            JSONObject json = new JSONObject();

            // Create JSON
            String configName = "weaponStuffs";
            JSONArray rows = ExcelParser.parseSheet(workbook, configName);
            json.put(configName, rows);

            String jsonText = json.toString();

            // Get the JSON text.
            Path path = Paths.get(fileName + "1.expected.json");
            BufferedReader reader = Files.newBufferedReader(path);
            String expected = reader.readLine();

            JSONAssert.assertEquals(expected, jsonText, false);
        }
    }

    @Test
    public void parseSheetWithArrays() throws Exception {
        String fileName = "testcases/test";
        File excelFile = new File(fileName + ".xlsx");

        try(FileInputStream inp = new FileInputStream( excelFile )) {
            Workbook workbook = WorkbookFactory.create(inp);
            // Start constructing JSON.
            JSONObject json = new JSONObject();

            // Create JSON
            String configName = "skillStuffs";
            JSONArray rows = ExcelParser.parseSheet(workbook, configName);
            json.put(configName, rows);

            String jsonText = json.toString();

            // Get the JSON text.
            Path path = Paths.get(fileName + "2.expected.json");
            BufferedReader reader = Files.newBufferedReader(path);
            String expected = reader.readLine();

            JSONAssert.assertEquals(expected, jsonText, false);
        }
    }

    @Test
    public void parseSheetWithObjectAndReference() throws Exception {
        String fileName = "testcases/test";
        File excelFile = new File(fileName + ".xlsx");

        try(FileInputStream inp = new FileInputStream( excelFile )) {
            Workbook workbook = WorkbookFactory.create(inp);
            // Start constructing JSON.
            JSONObject json = new JSONObject();

            // Create JSON
            String configName = "map";
            JSONArray rows = ExcelParser.parseSheet(workbook, configName);
            json.put(configName, rows);

            String jsonText = json.toString();

            // Get the JSON text.
            Path path = Paths.get(fileName + "3.expected.json");
            BufferedReader reader = Files.newBufferedReader(path);
            String expected = reader.readLine();

            JSONAssert.assertEquals(expected, jsonText, false);
        }
    }

}