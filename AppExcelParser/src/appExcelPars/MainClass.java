
package appExcelPars;

import java.io.IOException;


public class MainClass {
     public static void main(String[] args) throws IOException {
        System.out.println(Parser.parse("Person.xls"));
    }
}
