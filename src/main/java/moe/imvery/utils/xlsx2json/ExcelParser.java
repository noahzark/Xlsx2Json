package moe.imvery.utils.xlsx2json;

import org.apache.poi.ss.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by Feliciano on 6/1/2016.
 */
public class ExcelParser {

    public static String SIGN_HIDDEN_CELL_PREFIX = "$";

    public static String SIGN_ITEM_SPLITTER = ",";

    public static String SIGN_KEYVALUE_SPLITTER = ":";

    public static String SIGN_TABLE_REFERENCE_SPLITTER = "@";

    public static String SIGN_SHEETNAME_COLUMNNAME_SPLITTER = "#";


    /**
     * Parse the whole sheet of a workbook
     * @param workbook
     * @param configName
     * @return
     */
    public static JSONArray parseSheet( Workbook workbook, String configName ) {
        // Iterate through the rows.
        JSONArray rows = new JSONArray();

        // Get the Sheet by name.
        ParsedSheet parsedSheet = new ParsedSheet(workbook, configName);
        parsedSheet.parseSheet();

        Sheet sheet = parsedSheet.getSheet();

        // Parse each row.
        for (Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext(); )
        {
            Row row = rowsIT.next();

            if ( row.getRowNum() <= parsedSheet.nameRowIndex )
                continue;

            // Iterate through the cells.
            JSONObject jsonRow = parseRow(row, parsedSheet);

            rows.put( jsonRow );
        }

        return rows;
    }

    /**
     * Find a row using the key and value in another sheet
     * @param sheet The target sheet
     * @param key One column's name of the row
     * @param value The column's value
     * @return Found row, or null if not found
     */
    public static Row findRowByColumn( ParsedSheet sheet, String key, String value) {
        int index = sheet.indexOfKey(key);

        if (index == -1)
            throw new IllegalArgumentException("Couldn't find key " + key + " in the provided sheet.");

        for (Iterator<Row> rowsIT = sheet.getSheet().rowIterator(); rowsIT.hasNext(); ) {
            Row row = rowsIT.next();

            Cell cell = row.getCell(index);

            switch (sheet.getType(index)) {
                case BASIC:
                    if (cell == null)
                        continue;

                    if (cell.getCellType() == CellType.BLANK)
                        continue;

                    String cellValue = getCellStringValue(cell);

                    if (cellValue.equals(value))
                        return row;

                    break;

                default:
                    throw new IllegalArgumentException("Reference search doesn't support the type " + sheet.getType(index) + " of key " + key + ".");
            }
        }

        return null;
    }

    /**
     * Parse a row of the sheet
     * @param row The target row to parse
     * @param parsedSheet Parsed sheet to provide name and type information
     * @return A parsed JSONObject
     */
    public static JSONObject parseRow(Row row, ParsedSheet parsedSheet) {
        JSONObject jsonRow = new JSONObject();

        //Parse each cell
        for ( int index = 0; index < parsedSheet.width;  index++)
        {
            //System.out.println("Column: " + index);

            Cell cell = row.getCell(index);

            String key = parsedSheet.getKey( index );

            // Skip cells with empty key
            if (key.isEmpty())
                continue;

            // Skip hidden cells
            if (key.startsWith(SIGN_HIDDEN_CELL_PREFIX))
                continue;

            ParsedCellType type = parsedSheet.getType( index );

            // Null cell handler
            switch (type) {
                case BASIC:
                    // Handle "Null" string
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        if (cell.getStringCellValue().equalsIgnoreCase("null")) {
                            jsonRow.put(key, JSONObject.NULL);
                            continue;
                        }
                    }
                case DATE:
                case TIME:
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        if (cell.getStringCellValue().equalsIgnoreCase("null")) {
                            jsonRow.put(key, JSONObject.NULL);
                            continue;
                        }
                    }
                case OBJECT:
                    if (cell == null || cell.getCellType() == CellType.BLANK) {
                        jsonRow.put(key, JSONObject.NULL);
                        continue;
                    }
                    break;
                case REFERENCE:
                    if (cell == null || cell.getCellType() == CellType.BLANK) {
                        jsonRow.put(key.substring(0, key.indexOf("@")), JSONObject.NULL);
                        continue;
                    }
                    break;

                case ARRAY_STRING:
                case ARRAY_BOOLEAN:
                case ARRAY_DOUBLE:
                    if (cell == null || cell.getCellType() == CellType.BLANK) {
                        jsonRow.put(key, new ArrayList());
                        continue;
                    }
                    break;

                default:
                    throw new IllegalArgumentException("Unhandled empty cell of " + type + " type.");
            }

            String stringCellValue = "";

            switch (type) { // Add single value support for the row
                case ARRAY_STRING:
                case ARRAY_BOOLEAN:
                case ARRAY_DOUBLE:
                {
                    switch (cell.getCellType()) {
                        case BOOLEAN:
                            stringCellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case NUMERIC:
                            stringCellValue = String.valueOf(cell.getNumericCellValue());
                            break;
                        default:
                            stringCellValue = cell.getStringCellValue();
                    }
                }

                default:
            }

            switch (type) {
                case BASIC:
                    switch (cell.getCellType())
                    {
                        case NUMERIC:
                            jsonRow.put( key, cell.getNumericCellValue() );
                            break;
                        case BOOLEAN:
                            jsonRow.put( key, cell.getBooleanCellValue() );
                            break;
                        default:
                            jsonRow.put( key, cell.getStringCellValue() );
                            break;
                    };
                    break;

                case DATE:
                    if(cell.getCellType() == CellType.NUMERIC){
                        BigDecimal data = new BigDecimal(cell.getNumericCellValue()+"");
                        String dateStr = data+"";
                        Date date = null;
                        if ((date=isValidDate(dateStr)) !=null){
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                            dateStr = sdf.format(date);
                            jsonRow.put(key, dateStr);
                        } else{
                            jsonRow.put(key, "1990-01-01");
                        }
                    } else if(cell.getCellType() == CellType.STRING){
                        String dateStr = cell.getStringCellValue();
                        Date date = null;
                        if ((date=isValidDate(dateStr)) !=null){
                            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                            dateStr = sdf.format(date);
                            jsonRow.put(key, dateStr);
                        } else{
                            jsonRow.put(key, "1990-01-01");
                        }
                    } else {
                        throw new IllegalArgumentException("Unhandled cell of " + cell.getCellType()+ " type at "
                                + "row " + row.getRowNum()
                                + "column " + index);
                    }
                    break;

                case TIME:
                    if (cell.getCellType() == CellType.NUMERIC) {
                        Date time = cell.getDateCellValue();
                        Calendar calendar = Calendar.getInstance();
                        calendar.setTime(time);
                        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
                        String timeStr = sdf.format(calendar.getTime());
                        jsonRow.put( key, timeStr );
                    } else if (cell.getCellType() == CellType.STRING){
                        jsonRow.put( key, cell.getStringCellValue() );
                    } else {
                        throw new IllegalArgumentException("Unhandled cell of " + cell.getCellType()+ " type at "
                                + "row " + row.getRowNum()
                                + "column " + index);
                    }
                    break;

                case ARRAY_STRING:
                    ArrayList stringArray = ExcelParser.<ArrayList<String>>parseCellData(type, stringCellValue );
                    jsonRow.put( key, new JSONArray(stringArray) );
                    break;

                case ARRAY_BOOLEAN:
                    ArrayList booleanArray = ExcelParser.<ArrayList<Boolean>>parseCellData(type, stringCellValue );
                    jsonRow.put( key, new JSONArray(booleanArray) );
                    break;

                case ARRAY_DOUBLE:
                    ArrayList doubleArray = ExcelParser.<ArrayList<Double>>parseCellData(type, stringCellValue );
                    jsonRow.put( key, new JSONArray(doubleArray) );
                    break;

                case OBJECT:
                    JSONObject jsonObject = ExcelParser.parseCellData(type, cell.getStringCellValue());
                    jsonRow.put( key, jsonObject );
                    break;

                case REFERENCE:
                    // Split key to get real key, target sheet name and target column name
                    // Key example: monster@monsterSheet#monsterId
                    String[] keyAndTarget = key.split(SIGN_TABLE_REFERENCE_SPLITTER);
                    key = keyAndTarget[0];

                    // Split sheet name and column name
                    String[] realTarget = keyAndTarget[1].split(SIGN_SHEETNAME_COLUMNNAME_SPLITTER);
                    String targetSheetName = realTarget[0];
                    String targetKey = realTarget[1];
                    String targetValue = getCellStringValue(cell);

                    Sheet targetSheet = parsedSheet.getSheet(targetSheetName);
                    ParsedSheet parsedTargetSheet = new ParsedSheet(targetSheet.getWorkbook(), targetSheetName);
                    parsedTargetSheet.parseSheet();

                    Row targetRow = findRowByColumn(parsedTargetSheet, targetKey, targetValue);
                    jsonObject = parseRow(targetRow, parsedTargetSheet);

                    jsonRow.put( key, jsonObject);
                    break;

                default:
                    throw new IllegalArgumentException("Unsupported type " + type + " found.");
            }

        }
        return jsonRow;
    }

    private static String getCellStringValue(Cell cell) {
        switch (cell.getCellType()) {
            case BLANK:
                break;

            case NUMERIC:
                return cell.getNumericCellValue() + "";

            case BOOLEAN:
                return cell.getBooleanCellValue() + "";

            default:
                return cell.getStringCellValue();
        }

        return null;
    }

    private static Date isValidDate(String dateStr){
        Date date = null;
        SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");
        try{
            format.setLenient(false);
            date = format.parse(dateStr);
        }catch (ParseException e){
            return null;
        }
        return date;
    }

    /**
     * Parse a cell of the row
     * @param type The data type
     * @param cellValue The cell string to be parsed
     * @param <W> The return data type
     * @return Parsed data
     * @throws NumberFormatException Numeric data parse failed
     */
    public static <W> W parseCellData(ParsedCellType type, String cellValue) throws NumberFormatException {
        Object object = null;

        String[] items = cellValue.split(SIGN_ITEM_SPLITTER);

        switch (type) {
            case ARRAY_STRING:
                ArrayList<String> arrayString = new ArrayList<>();
                for (String item : items) {
                    item = item.trim();
                    arrayString.add(item);
                }
                object = arrayString;
                break;

            case ARRAY_BOOLEAN:
                ArrayList<Boolean> arrayBoolean = new ArrayList<>();
                for (String item : items) {
                    item = item.trim();
                    arrayBoolean.add(Boolean.parseBoolean(item));
                }
                object = arrayBoolean;
                break;

            case ARRAY_DOUBLE:
                ArrayList<Double> arrayDouble = new ArrayList<>();
                for (String item : items) {
                    item = item.trim();
                    arrayDouble.add(Double.parseDouble(item));
                }
                object = arrayDouble;
                break;

            case OBJECT:
                JSONObject obj = new JSONObject();

                for (String item : items) {
                    String temp = item.trim();

                    String[] keyValue = item.split(SIGN_KEYVALUE_SPLITTER);
                    String key = keyValue[0], value = keyValue[1];
                    key = key.trim();
                    value = value.trim();

                    // Handle the null child
                    if (value.equalsIgnoreCase("null")) {
                        obj.put( key, JSONObject.NULL );
                        continue;
                    }

                    if (value.startsWith("\"")) {
                        obj.put( key, value.substring(1, value.length()-1));
                        continue;
                    }

                    try {
                        obj.put ( key, Double.parseDouble(value) );
                    } catch (NumberFormatException e) {
                        if (Boolean.parseBoolean(value)) {
                            obj.put( key, true );
                        } else if (value.equalsIgnoreCase("false")) {
                            obj.put( key, false);
                        } else {
                            obj.put( key, value);
                        }
                    }
                }

                object = obj;
                break;
        }

        return (W) object;
    }

}
