package pars.excel.doc;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

public class PersonPars {

    public static final int NAME_COLUMN_NUMBER = 0;
    public static final int ADDRESS_COLUMN_NUMBER = 1;
    public static final int PHONE_COLUMN_NUMBER = 2;

    public static void readFromExcel(String fileName) throws IOException {
        HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(fileName));
        HSSFSheet sheet = workBook.getSheetAt(0);

        Iterator<Row> rows = sheet.rowIterator();

        if (rows.hasNext()) {
            rows.next();
        }
        while (rows.hasNext()) {
            HSSFRow row = (HSSFRow) rows.next();
            //получаем значение ячеек по номерам столбцов
            HSSFCell nameCell = row.getCell(NAME_COLUMN_NUMBER);
            //получаем строковое значение из ячейки
            String name = nameCell.getStringCellValue();
            System.out.println("Имя: " + name);
            HSSFCell addressCell = row.getCell(ADDRESS_COLUMN_NUMBER);
            String address = addressCell.getStringCellValue();
            System.out.println("Адрес: " + address);
            HSSFCell phoneNumberCell = row.getCell(PHONE_COLUMN_NUMBER);
            String phoneNumber = phoneNumberCell.getStringCellValue();
            System.out.println("Номер телефона: " + phoneNumber);
        }
    }
}
