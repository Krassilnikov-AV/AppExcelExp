package pars.excel.doc;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class PersonParsList {

    public static void readFirstRow(String filename) {
        List<String> fieldsArrayList = new ArrayList<String>();
        try {
            FileInputStream myInput = new FileInputStream(new File(filename));

            Workbook workBook = null;
            workBook = new HSSFWorkbook(myInput);
            Sheet sheet = workBook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);

            int length = firstRow.getLastCellNum();

            Cell cell = null;
            for (int i = 0; i < length; i++) {
                cell = firstRow.getCell(i);
                fieldsArrayList.add(cell.toString());
           }
            Iterator ir=fieldsArrayList.iterator();
  while(ir.hasNext()){
System.out.println(ir.next());
  }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public void readColumnValues(int cityPosition, String fileName) { 
 try  { 
  FileInputStream myInput = new FileInputStream(fileName); 

  Workbook workBook = null; 

  workBook = new HSSFWorkbook(myInput); 
  Sheet sheet = workBook.getSheetAt(0); 

  for (int i = 0 ; i <= sheet.getLastRowNum() ; i++)   { 
   Row row = sheet.getRow(i); 

   if (i > 0) //skip first row 
   { 
    Cell cityCell = row.getCell(cityPosition); 
    String cityNames  = cityCell.toString(); 
   } 
  } 
 }
 catch(Exception e)  { 
  e.printStackTrace(); 
 } 
} 
}
