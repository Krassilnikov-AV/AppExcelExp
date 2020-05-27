package appReadData;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class ReadExcelFaile {

    public static final int ID_COLUMN_NUMBER = 0;
    public static final int NAME_COLUMN_ASIAN = 1;
    public static final int NAME_COLUMN_EUROPA = 2;

    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("123.xls");
        Workbook wb = new HSSFWorkbook(fis);
        HSSFSheet sheet = (HSSFSheet) wb.getSheetAt(0);

        Iterator<Row> rows = sheet.rowIterator();
        if (rows.hasNext()) {
            rows.next();
        }
        while (rows.hasNext()) {
            HSSFRow row = (HSSFRow) rows.next();
            //получаем значение ячеек по номерам столбцов
            HSSFCell idCell = row.getCell(ID_COLUMN_NUMBER);
            //получаем строковое значение из ячейки
            String id = idCell.getStringCellValue();
            System.out.println("ID: " + id);
            HSSFCell asianCell = row.getCell(NAME_COLUMN_ASIAN);
            String asian = asianCell.getStringCellValue();
            System.out.println("Азия: " + asian);
            HSSFCell europa = row.getCell(NAME_COLUMN_EUROPA);
            String phoneNumber = europa.getStringCellValue();
            System.out.println("Европа: " + phoneNumber);
        }

//выбор листа _ строки _ столбца
// преобразование с вновь введеным методом, который определит формат считываемого знвчения
//        String result0 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(1));
//        String result1 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(3));
//        String result2 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(4));
//        String result3 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(5));
//        String result4 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(6));
//        String result5 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(2));
//        System.out.println(result0 + "->" + result1 + "->" + result2 + "->" + result3 + "->" + result4 + "->" + result5);
    }

// метод для самостоятельного определения формата считываемого значения
    public static String getCelltext(HSSFCell cell) {

        String result = "";

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = cell.getDateCellValue().toString();
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = cell.getCellFormula().toString();
                break;
            default:
                break;
        }
        return result;
    }
}
