package ru.write.word.example;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 *
 * @author Aleks
 */
public class CreatTableClass {

    /**
     * Класс позволяет создавать таблицу в Word документе
     *
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, InvalidFormatException {
//        XWPFDocument doc = new XWPFDocument();
//        XWPFParagraph parag = doc.createParagraph();
//        XWPFRun run = parag.createRun();
//        run.setText("Очевидно и не вероятно!");
//
//        String folder = "D:/temp/";
//        String fileName = "nwTable";
//// создание папки с существующим файлом, если нет такой на диске         
//        File f = new File(folder);
//        if (!f.exists()) {
//            System.out.println("Created folder " + folder);
//            f.mkdirs();
//        }
//
//        FileOutputStream out = new FileOutputStream(new File(folder + fileName));
//        doc.write(out);
//        //      doc.close();
//
//        System.out.println("Записан файл: " + folder + fileName);
//        out.close();
        
        createDocFormTemplate();
    }

    private static void createDocFormTemplate() throws InvalidFormatException, IOException {
        XWPFDocument doc = new XWPFDocument();
        List<String> contents = new ArrayList<String>();
        contents.add("John");
        contents.add("Maikle");
        contents.add("Sanya");
        contents.add("Nike");

        XWPFTable table = doc.getTables().get(0);
        for (int i = 0; i < contents.size(); i++) {
            table.createRow().getCell(0).setText(contents.get(i));
        }

        write2File(doc);
    }

    private static void write2File(XWPFDocument doc) throws FileNotFoundException, IOException {
        String folder = "D:/temp/";
        String fileName = "createFileFromTemplate.docx";

        File f = new File(folder);
        if (!f.exists()) {
            System.out.println("Created folder " + folder);
            f.mkdirs();
        }

        FileOutputStream out = new FileOutputStream(new File(folder + fileName));
        doc.write(out);
    }

    /**
     * - код для разбора - для Word 95 private static void test(int rows, int
     * columns) throws Exception { // POI apparently can't create a document
     * from scratch, // so we need an existing empty dummy document
     * POIFSFileSystem fs = new POIFSFileSystem(new
     * FileInputStream("empty.doc")); XWPFDocument doc = new XWPFDocument(fs);
     *
     * Range range = doc.getRange(); Table table = range.insertBefore(new
     * TableProperties(columns), rows);
     *
     * for (int rowIdx = 0; rowIdx < table.numRows(); rowIdx++) { TableRow row =
     * table.getRow(rowIdx); System.out.println("row " + rowIdx); for (int
     * colIdx = 0; colIdx < row.numCells(); colIdx++) { TableCell cell =
     * row.getCell(colIdx); System.out.println("column " + colIdx + ", num
     * paragraphs " + cell.numParagraphs()); try { Paragraph par =
     * cell.getParagraph(0); par.insertBefore("" + (rowIdx * row.numCells() +
     * colIdx)); } catch (Exception ex) { ex.printStackTrace(); } } } }
     */
}
