package appWriteData;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class WriteExacelFaile {

    public static void main(String[] args) throws FileNotFoundException, IOException {
// создание /открытие книги - работа с книгой
        Workbook wb = new HSSFWorkbook();
// Создайте лист для этой рабочей книги, добавьте его на листы и
// верните представление высокого уровня.
        Sheet sheet0 = wb.createSheet("Сontinent-Сountry");
        Row row = sheet0.createRow(1); // создание строки

        Cell cell00 = row.createCell(0); // создание столбца
        cell00.setCellValue("id_"); // запись в ячейку

        Cell cell0 = row.createCell(1); // создание столбца
        cell0.setCellValue("Asian"); // запись в ячейку

        Cell cell1 = row.createCell(2);
        cell1.setCellValue("Europa");

        Cell cell2 = row.createCell(3);
        cell2.setCellValue("Africa");

        Cell cell3 = row.createCell(4);
        cell3.setCellValue("Southern America");

        Cell cell4 = row.createCell(5);
        cell4.setCellValue("Northern America");

        Cell cell5 = row.createCell(6);
        cell5.setCellValue("Australia");

        Sheet sheet1 = wb.createSheet("city");

        Sheet sheet2 = wb.createSheet("products");
// создание страницы со всякой белебердой     
        Sheet sheet3 = wb.createSheet(WorkbookUtil.createSafeSheetName("sdvgsert@#^&*&&"));
// создали поток файла
        FileOutputStream fos = new FileOutputStream("123.xls");

// записали книгу в файл
        wb.write(fos);
// закрыли поток
        fos.close();
// =========> файл и  листы созданы, записи добавлены <==========
    }
}
