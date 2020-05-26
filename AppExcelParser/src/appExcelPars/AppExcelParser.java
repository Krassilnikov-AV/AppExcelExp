
package appExcelPars;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
//import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

/**
 *Обработка файлов .xls с помощью Apache POI
 * В ходе работы над различными программными продуктами часто возникает необходимость 
 * импорта и экспорта данных из различных "закрытых" форматов файлов. Чаще всего эта
 * необходимость возникает применительно к файлам в форматах офисных продуктов корпорации Microsoft,
 * в частности Word (doc, docx) и Excel (xls, xlsx). В силу особенностей реализации этих форматов, 
 * реализация такой обработки "в лоб" целиком своими силами была бы весьма нетривиальной и достаточно 
 * трудоёмкой задачей. К счастью, основная часть работы уже сделана за нас - существуют открытые Java-библиотеки,
 * позволяющие преобразовать эти файлы в объектную модель Java, после чего получение из них необходимой нам 
 * информации не составит особого труда. В этой заметке показан пример реализации такой выборки данных из документа 
 * .xls с помощью Apache POI (на примере версии 3.6) - популярной библиотеки от Apache Software Foundation.
 */

public class AppExcelParser {

public static final NAME_COLUMN_NUMBER = 0; //ФИО
public static final ADDRESS_COLUMN_NUMBER = 1; //Адрес
public static final PHONE_COLUMN_NUMBER = 2; //Телефон
    
    public static void main(String[] args) {
       
    }
    
    public List<ContactPerson> getContacts(String path) throws IOException{
        List<ContactPerson> contacts = new ArrayList<ContactPerson>(); //Создаём пустой список контактов
 
        File addressDB = new File(path); //Переменная path содержит путь к документу в ФС
        POIFSFileSystem fileSystem = new POIFSFileSystem(addressDB); //Открываем документ
        HSSFWorkbook workBook = new HSSFWorkbook(fileSystem); // Получаем workbook
        HSSFSheet sheet = workBook.getSheetAt(0); // Проверяем только первую страницу
 
        Iterator<Row> rows = sheet.rowIterator(); // Перебираем все строки
 
        // Пропускаем "шапку" таблицы
        if (rows.hasNext()) {
                rows.next();
        }
 
        // Перебираем все строки начиная со второй до тех пор, пока документ не закончится 
        while (rows.hasNext()) {
                HSSFRow row = (HSSFRow) rows.next();
                //Получаем ячейки из строки по номерам столбцов
                HSSFCell nameCell = row.getCell(NAME_COLUMN_NUMBER); //ФИО
                HSSFCell addressCell = row.getCell(ADDRESS_COLUMN_NUMBER); //Адрес
                HSSFCell phoneCell = row.getCell(PHONE_COLUMN_NUMBER); //Номер телефона
                // Если в первом столбце нет данных, то контакт не создаём 
                if (nameCell != null) {
                        ContactPerson person = new ContactPerson();
                        person.setName(nameCell.getStringCellValue()); //Получаем строковое значение из ячейки
 
                        person.setAddress(""); //Адрес может не быть задан
                        if (addressCell != null && !"".equals(addressCell.getStringCellValue())) {
                                person.setAddress(addressCell.getStringCellValue()); //Адрес - строка
                        }
 
                        person.setPhone(""); //Телефон тоже может не быть задан
                        if (phoneCell != null && !"".equals(phoneCell.getStringCellValue())) {
                                person.setPhoneNumber(phoneCell.getStringCellValue()); // Телефон - тоже строка
                        }
 
                        contacts.add(person); //Добавляем контакт в список
 
                }
        }
        return contacts;
}
}
