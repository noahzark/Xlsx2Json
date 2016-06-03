package moe.imvery.utils.xlsx2json;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.skyscreamer.jsonassert.JSONAssert;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Created by Feliciano on 6/1/2016.
 */
public class ExcelParserMain {

    /**
     * Parse the excel file and save as json
     * @param targetName Excel file name without suffix
     * @param sheetList Target sheet list
     * @param showSheetName Whether show sheet name in result or not
     */
    public static void parseExcel(String targetName, String[] sheetList, boolean showSheetName) {
        File excelFile = new File(targetName + ".xlsx");

        try {
            FileInputStream inp = new FileInputStream( excelFile );
            Workbook workbook = WorkbookFactory.create( inp );

            String jsonText;

            if (showSheetName) {
                // Start constructing JSON.
                JSONObject json = new JSONObject();

                // Create JSON
                for (String sheetName : sheetList) {
                    JSONArray rows = ExcelParser.parseSheet(workbook, sheetName);
                    json.put(sheetName, rows);
                }

                // Get the JSON text.
                jsonText = json.toString();
            } else {
                JSONArray json = new JSONArray();

                // Create JSON
                for (String sheetName : sheetList) {
                    JSONArray rows = ExcelParser.parseSheet(workbook, sheetName);
                    for (int i=0; i < rows.length(); i++) {
                        json.put(rows.get(i));
                    }
                }

                // Get the JSON text.
                jsonText = json.toString();
            }

            // Write into file
            Path path = Paths.get(targetName + ".json");
            BufferedWriter writer = Files.newBufferedWriter( path );
            writer.write(jsonText);
            writer.close();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Validate two json file.  If it isn't it throws an {@link AssertionError}.
     * @param expectedFileName Expected JSON file
     * @param targetFileName File to compare
     * @param strict Enables strict checking
     * @throws JSONException
     */
    public static void validateJson(String expectedFileName, String targetFileName, boolean strict) {
        try {
            System.out.println("Checking " + targetFileName + " ... ");
            Path path = Paths.get(targetFileName);
            BufferedReader reader = Files.newBufferedReader(path);
            String jsonText = reader.readLine();

            path = Paths.get(expectedFileName);
            reader = Files.newBufferedReader(path);
            String expected = reader.readLine();

            JSONAssert.assertEquals(expected, jsonText, false);

            System.out.println("Passed.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        if (args.length < 2)
            throw new IllegalArgumentException("Expected at least 2 arguments representing Filename, Sheetnames(Divided by space, surrounded by \").");

        if (args.length > 3)
            throw new IllegalArgumentException("Expected at most 3 arguments representing Filename, Sheetnames(Divided by space, surrounded by \") and a boolean for show sheet names in result or not.");

        String targetName = args[0];

        if (targetName.equalsIgnoreCase("Validate")) {

        } else {
            if (!targetName.endsWith("xlsx"))
                throw new IllegalArgumentException("The first argument should be a excel(xlsx) file name.");
            // Cut the .xlsx suffix
            targetName = targetName.substring(0, targetName.length()-5);

            // Split sheet names
            String[] sheetList = args[1].split(" ");

            // Detect show sheet name option
            boolean showSheetName = (args.length == 3) ? Boolean.parseBoolean(args[2]) : false;

            parseExcel(targetName, sheetList, showSheetName);
        }
    }

}
