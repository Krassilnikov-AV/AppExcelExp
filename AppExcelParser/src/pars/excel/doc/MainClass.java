
package pars.excel.doc;

import java.io.IOException;

public class MainClass {
   
    public static void main(String[] args) throws IOException {
         String fileName = "Person.xls";
  //      PersonPars.readFromExcel(fileName);    
        PersonParsList.readFirstRow(fileName);
    }    
}
