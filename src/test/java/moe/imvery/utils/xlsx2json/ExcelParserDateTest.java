package moe.imvery.utils.xlsx2json;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.junit.Test;

import java.util.Calendar;
import java.util.GregorianCalendar;

import static org.junit.Assert.assertEquals;

public class ExcelParserDateTest {

    private static final String SHEET = "dates";

    /**
     * Create a workbook with a single Date column holding the given value
     */
    private static Workbook workbookWithDate(Object value) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(SHEET);

        Row typeRow = sheet.createRow(0);
        typeRow.createCell(0).setCellValue("Basic");
        typeRow.createCell(1).setCellValue("Date");

        Row nameRow = sheet.createRow(1);
        nameRow.createCell(0).setCellValue("id");
        nameRow.createCell(1).setCellValue("date");

        Row row = sheet.createRow(2);
        row.createCell(0).setCellValue(1);
        Cell cell = row.createCell(1);
        if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Calendar) {
            CellStyle style = workbook.createCellStyle();
            style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("yyyy-mm-dd"));
            cell.setCellStyle(style);
            cell.setCellValue((Calendar) value);
        } else {
            cell.setCellValue((String) value);
        }

        return workbook;
    }

    private static String parseDate(Object value) {
        JSONArray rows = ExcelParser.parseSheet(workbookWithDate(value), SHEET);
        return rows.getJSONObject(0).getString("date");
    }

    @Test
    public void numericDate() {
        assertEquals("2016-06-01", parseDate(20160601));
    }

    @Test
    public void numericDateEndingInZero() {
        // These used to fall back to 1990-01-01, because Double.toString() gave "2.016101E7"
        assertEquals("2016-10-10", parseDate(20161010));
        assertEquals("2016-11-20", parseDate(20161120));
        assertEquals("2016-12-30", parseDate(20161230));
    }

    @Test
    public void stringDate() {
        assertEquals("2016-10-10", parseDate("20161010"));
    }

    @Test
    public void excelDateCell() {
        assertEquals("2016-10-10", parseDate(new GregorianCalendar(2016, Calendar.OCTOBER, 10)));
    }

    @Test(expected = IllegalArgumentException.class)
    public void invalidMonthIsRejected() {
        parseDate("20161310");
    }

    @Test(expected = IllegalArgumentException.class)
    public void trailingCharactersAreRejected() {
        parseDate(201610105);
    }

    @Test(expected = IllegalArgumentException.class)
    public void fractionalNumberIsRejected() {
        parseDate(20161010.5);
    }
}
