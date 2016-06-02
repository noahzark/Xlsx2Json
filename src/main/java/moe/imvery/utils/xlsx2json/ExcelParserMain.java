package moe.imvery.utils.xlsx2json;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Created by Feliciano on 6/1/2016.
 */
public class ExcelParserMain {

    public static void main(String[] args) {
        String fileName = "test";
        File excelFile = new File(fileName + ".xlsx");

        try(FileInputStream inp = new FileInputStream( excelFile )) {
            Workbook workbook = WorkbookFactory.create( inp );
            // Start constructing JSON.
            JSONObject json = new JSONObject();

            // Create JSON
            String configName = "shieldStuffs";
            JSONArray rows = ExcelParser.parseSheet(workbook, configName);
            json.put( configName, rows );

            // Get the JSON text.
            String jsonText = json.toString();
            Path path = Paths.get(fileName + ".json");
            BufferedWriter writer = Files.newBufferedWriter( path );
            writer.write(jsonText);
            writer.close();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
