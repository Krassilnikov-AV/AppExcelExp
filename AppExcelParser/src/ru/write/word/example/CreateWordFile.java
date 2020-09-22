package ru.write.word.example;

import java.io.IOException;
import java.math.BigInteger;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

/**
 *
 * @author Aleks
 */
public class CreateWordFile {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, XmlException {
        // создаем модель docx документа, 
        // к которой будем прикручивать наполнение (колонтитулы, текст)
        XWPFDocument docxModel = new XWPFDocument();

        // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxModel);
    }

    /* Добавить
    // SBT config
"org.apache.poi" % "poi-ooxml" % "4.1.0",     // Base library
"org.apache.poi" % "ooxml-schemas" % "1.4",   // required to access CTPageSz
     */
    private void changeOrientation(XWPFDocument document, String orientation) {
        CTDocument1 doc = document.getDocument();
        CTBody body = doc.addNewBody();
        body.addNewSectPr();
        CTSectPr section = body.getSectPr();
        if (!section.isSetPgSz()) {
            section.addNewPgSz();
        }
        CTPageSz pageSize = section.getPgSz();
 //       CTPageSz pageSize;
        if (section.isSetPgSz()) {
            pageSize = section.getPgSz();
        } else {
            pageSize = section.addNewPgSz();
        }
        if (orientation.equals("landscape")) {
            pageSize.setOrient(STPageOrientation.LANDSCAPE);
            pageSize.setW(BigInteger.valueOf(842 * 20));
            pageSize.setH(BigInteger.valueOf(595 * 20));
        } else {
            pageSize.setOrient(STPageOrientation.PORTRAIT);
            pageSize.setH(BigInteger.valueOf(842 * 20));
            pageSize.setW(BigInteger.valueOf(595 * 20));
        }
    }
/**
    private void changeOrientation(CTSectPr section, String orientation) {
        CTPageSz pageSize = section.isSetPgSz ? section.getPgSz() : section.addNewPgSz();
        if (orientation.equals("landscape")) {
            pageSize.setOrient(STPageOrientation.LANDSCAPE);
            pageSize.setW(BigInteger.valueOf(842 * 20));
            pageSize.setH(BigInteger.valueOf(595 * 20));
        } else {
            pageSize.setOrient(STPageOrientation.PORTRAIT);
            pageSize.setH(BigInteger.valueOf(842 * 20));
            pageSize.setW(BigInteger.valueOf(595 * 20));
        }
    }
    * */
}
