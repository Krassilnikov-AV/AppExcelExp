package appReadData;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;

public class ReadExcelFaile {

    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream("123.xls");
        Workbook wb = new HSSFWorkbook(fis);
        //выбор листа _ строки _ столбца
//        String result0 = wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue();
//        String result1 = wb.getSheetAt(0).getRow(1).getCell(2).getStringCellValue();
//        String result2 = wb.getSheetAt(0).getRow(1).getCell(3).getStringCellValue();
//        String result3 = wb.getSheetAt(0).getRow(1).getCell(4).getStringCellValue();
//        String result4 = wb.getSheetAt(0).getRow(1).getCell(5).getStringCellValue();
//        String result5 = wb.getSheetAt(0).getRow(1).getCell(6).getStringCellValue();
// преобразование с вновь введеным методом, который определит формат считываемого знвчения
        String result0 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(1));
        String result1 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(3));
        String result2 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(4));
        String result3 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(5));
        String result4 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(6));
        String result5 = getCelltext(wb.getSheetAt(0).getRow(1).getCell(2));

        System.out.println(result0 + "->" + result1 + "->" + result2 + "->" + result3 + "->" + result4 + "->" + result5);
    }
    
// метод для самостоятельного определения формата считываемого значения
    public static String getCelltext(Cell cell) {

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
