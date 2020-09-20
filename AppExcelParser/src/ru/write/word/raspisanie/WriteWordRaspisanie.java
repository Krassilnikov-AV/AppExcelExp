package ru.write.word.raspisanie;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * пример создания Word документа
 * @author Aleks
 */
public class WriteWordRaspisanie {
    
   public static void main(String[] args)throws Exception {
       try {
           XWPFDocument document = new XWPFDocument();
           FileOutputStream out = new FileOutputStream(new File("D:\\LEARNING\\JAVA_DEVELOP\\MyProject\\Repositories\\ApplicExcelPars\\poidemo.docx"));
           
           XWPFParagraph paragraph = document.createParagraph();
           XWPFRun run = paragraph.createRun();
           run.setText("Hello! My Word!");
           document.write(out);
           out.close();
       } catch(Exception e) {
           System.out.println("Документ не создан");
           System.out.println(e);
       }
       System.out.println("Документ создан");
   }
}
