package moe.imvery.utils.xlsx2json;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONArray;
import org.json.JSONObject;

import java.util.ArrayList;
import java.util.Iterator;

/**
 * Created by Feliciano on 6/1/2016.
 */
public class ConfigParser {

    /**
     * Parse the whole sheet of a workbook
     * @param workbook
     * @param configName
     * @return
     */
    public static JSONArray parseSheet(Workbook workbook, String configName) {
        // Iterate through the rows.
        JSONArray rows = new JSONArray();

        // Get the Sheet by name.
        Sheet sheet = workbook.getSheet(configName);

        int typeRowIndex = 0, nameRowIndex = 1;

        // Fetch the type row and the name row.
        ArrayList<ConfigType> types = new ArrayList<>();
        if ( !sheet.getRow(typeRowIndex).getCell(0).getStringCellValue().equalsIgnoreCase("basic") ) {
            // If the primary key doesn't have a type defined "Basic", then we'll think all the columns are basic type,
            // and the first row is name row.
            typeRowIndex = 0;
            nameRowIndex = 0;

            Row typeRow = sheet.getRow(typeRowIndex);
            for (Iterator<Cell> cellsIT = typeRow.cellIterator(); cellsIT.hasNext(); )
            {
                Cell cell = cellsIT.next();

                types.add(ConfigType.BASIC);
            }
        } else {
            // Else read the type of each column
            Row typeRow = sheet.getRow(typeRowIndex);
            for (Iterator<Cell> cellsIT = typeRow.cellIterator(); cellsIT.hasNext(); )
            {
                Cell cell = cellsIT.next();
                String cellType = cell.getStringCellValue();
                types.add(ConfigType.fromString(cellType));
            }
        }

        ArrayList<String> keys = new ArrayList<>();
        Row nameRow = sheet.getRow(nameRowIndex);
        for (Iterator<Cell> cellsIT = nameRow.cellIterator(); cellsIT.hasNext(); )
        {
            Cell cell = cellsIT.next();
            keys.add(cell.getStringCellValue());
        }

        for (Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext(); )
        {
            Row row = rowsIT.next();

            if ( row.getRowNum() <= nameRowIndex )
                continue;

            // Iterate through the cells.
            JSONObject jsonRow = parseRow(row, keys, types);

            rows.put( jsonRow );
        }

        return rows;
    }

    /**
     * Parse a row of the sheet
     * @param row
     * @param keys
     * @param types
     * @return
     */
    public static JSONObject parseRow(Row row, ArrayList<String> keys, ArrayList<ConfigType> types) {
        JSONObject jsonRow = new JSONObject();
        for ( Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext(); )
        {
            Cell value = cellsIT.next();
            int index = value.getColumnIndex();
            String key = keys.get( index );
            ConfigType type = types.get( index );

            ArrayList result;
            JSONArray jsonArray;

            switch (type) {
                case BASIC:
                    switch (value.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                            jsonRow.put( key, value.getNumericCellValue() );
                            break;
                        case Cell.CELL_TYPE_BOOLEAN:
                            jsonRow.put( key, value.getBooleanCellValue() );
                            break;
                        default:
                            jsonRow.put( key, value.getStringCellValue() );
                            break;
                    };
                    break;

                case ARRAY_STRING:
                    result = ConfigParser.<ArrayList<String>>parseCellData(type, value);
                    jsonArray = new JSONArray(result);
                    jsonRow.put( key, jsonArray );
                    break;

                case ARRAY_BOOLEAN:
                    result = ConfigParser.<ArrayList<Boolean>>parseCellData(type, value);
                    jsonArray = new JSONArray(result);
                    jsonRow.put( key, jsonArray );
                    break;

                case ARRAY_DOUBLE:
                    result = ConfigParser.<ArrayList<Double>>parseCellData(type, value);
                    jsonArray = new JSONArray(result);
                    jsonRow.put( key, jsonArray );
                    break;

                case REFERENCE:
                    // TODO Deal with the reference type
                    break;

                default:
                    throw new IllegalArgumentException("Unsupported type " + type + " found");
            }

        }
        return jsonRow;
    }

    /**
     * Parse a cell of the row
     * @param type
     * @param cell
     * @param <W>
     * @return
     * @throws NumberFormatException
     */
    public static <W> W parseCellData(ConfigType type, Cell cell) throws NumberFormatException {
        Object object = null;

        String cellValue = cell.getStringCellValue();
        String[] items;

        switch (type) {
            case ARRAY_STRING:
                items = cellValue.split(",");
                ArrayList<String> arrayString = new ArrayList<>();
                for (String item : items) {
                    item = item.trim();
                    arrayString.add(item);
                }
                object = arrayString;
                break;

            case ARRAY_BOOLEAN:
                items = cellValue.split(",");
                ArrayList<Boolean> arrayBoolean = new ArrayList<>();
                for (String item : items) {
                    item = item.trim();
                    arrayBoolean.add(Boolean.parseBoolean(item));
                }
                object = arrayBoolean;
                break;

            case ARRAY_DOUBLE:
                items = cellValue.split(",");
                ArrayList<Double> arrayDouble = new ArrayList<>();
                for (String item : items) {
                    item = item.trim();
                    arrayDouble.add(Double.parseDouble(item));
                }
                object = arrayDouble;
                break;
        }

        return (W) object;
    }

}
