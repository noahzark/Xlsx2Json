package moe.imvery.utils.xlsx2json;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Iterator;

/**
 * Created by Feliciano on 6/2/2016.
 */
public class ParsedSheet {

    private Workbook workbook;
    private Sheet sheet;

    private ArrayList<ParsedCellType> types;
    private ArrayList<String> keys;

    public int typeRowIndex, nameRowIndex;
    public int width;

    private boolean parsed;

    public ParsedSheet(Workbook workbook, String sheetName) {
        this.workbook = workbook;
        sheet = workbook.getSheet(sheetName);

        if (sheet == null)
            throw new IllegalArgumentException("Unable to find the sheet name " + sheetName + " in the workbook.");

        typeRowIndex = 0;
        nameRowIndex = 1;

        width = 0;

        parsed = false;

        types = new ArrayList<>();
        keys = new ArrayList<>();
    }

    public ParsedSheet parseSheet() {
        if (parsed)
            return this;

        try {
            // Fetch the type row.
            String firstColumnValue = sheet.getRow(typeRowIndex).getCell(0).getStringCellValue();
            if ( !ParsedCellType.isBasicType(firstColumnValue)) {
                // If the primary key doesn't have a type defined "Basic", then we'll think all the columns are basic type,
                // and the first row is name row.
                typeRowIndex = 0;
                nameRowIndex = 0;

                Row typeRow = sheet.getRow(typeRowIndex);
                for (Iterator<Cell> cellsIT = typeRow.cellIterator(); cellsIT.hasNext(); )
                {
                    Cell cell = cellsIT.next();
                    types.add(ParsedCellType.BASIC);
                }
            } else {
                // Else read the type of each column
                Row typeRow = sheet.getRow(typeRowIndex);
                for (Iterator<Cell> cellsIT = typeRow.cellIterator(); cellsIT.hasNext(); )
                {
                    Cell cell = cellsIT.next();
                    String cellType = cell.getStringCellValue();
                    types.add(ParsedCellType.fromString(cellType));
                }
            }

            // Fetch the name row.
            Row nameRow = sheet.getRow(nameRowIndex);
            for (Iterator<Cell> cellsIT = nameRow.cellIterator(); cellsIT.hasNext(); )
            {
                Cell cell = cellsIT.next();
                keys.add(cell.getStringCellValue());

                width++;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        parsed = true;

        return this;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Sheet getSheet(String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null)
            throw new IllegalArgumentException("Unable to find the sheet name " + sheetName + " in the workbook.");
        return sheet;
    }

    public boolean isParsed() {
        return parsed;
    }

    public ParsedCellType getType(int index) {
        if (!isParsed())
            throw new NullPointerException("This sheet haven't been parsed, please call parseSheet() method first!");

        return types.get(index);
    }

    public String getKey(int index) {
        if (!isParsed())
            throw new NullPointerException("This sheet haven't been parsed, please call parseSheet() method first!");

        return keys.get(index);
    }

    public int indexOfKey(String key) {
        if (!isParsed())
            throw new NullPointerException("This sheet haven't been parsed, please call parseSheet() method first!");

        return keys.indexOf(key);
    }

}
