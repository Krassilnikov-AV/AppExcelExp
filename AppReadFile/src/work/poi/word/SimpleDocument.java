
package work.poi.word;

import java.io.FileOutputStream; 
import java.io.IOException;

import org.apache.poi.xwpf.usermodel.Borders; 
import org.apache.poi.xwpf.usermodel.BreakClear; 
import org.apache.poi.xwpf.usermodel.BreakType; 
import org.apache.poi.xwpf.usermodel.LineSpacingRule; 
import org.apache.poi.xwpf.usermodel.ParagraphAlignment; 
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns; 
import org.apache.poi.xwpf.usermodel.VerticalAlign; 
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Простой документ WOrdprocessingML, созданный POI XWPF API
 */
public class SimpleDocument {

    public static void main (String [] args) throws IOException {
        try (XWPFDocument doc = new XWPFDocument ()) {

            XWPFParagraph p1 = doc.createParagraph ();
            p1.setAlignment (ParagraphAlignment.CENTER);
            p1.setBorderBottom (Borders.DOUBLE);
            p1.setBorderTop (Borders.DOUBLE);

            p1.setBorderRight (Borders.DOUBLE);
            p1.setBorderLeft (Borders.DOUBLE);
            p1.setBorderBetween (Borders.SINGLE);

            p1.setVerticalAlignment (TextAlignment.TOP);

            XWPFRun r1 = p1.createRun ();
            r1.setBold (true);
            r1.setText ("Быстрая коричневая лисица");
            r1.setBold (false);
            r1.setFontFamily ( "Courier");
            r1.setUnderline (UnderlinePatterns.DOT_DOT_DASH);
            r1.setTextPosition (100);

            XWPFParagraph p2 = doc.createParagraph ();
            p2.setAlignment (ParagraphAlignment.RIGHT);

            // ГРАНИЦ
            p2.setBorderBottom (Borders.DOUBLE);
            p2.setBorderTop (Borders.DOUBLE);
            p2.setBorderRight (Borders.DOUBLE);
            p2.setBorderLeft (Borders.DOUBLE);
            p2.setBorderBetween (Borders.SINGLE);

            XWPFRun r2 = p2.createRun ();
            r2.setText ("перепрыгнул через ленивую собаку");
            r2.setStrikeThrough (true);
            r2.setFontSize (20);

            XWPFRun r3 = p2.createRun ();
            r3.setText ("и ушел");
            r3.setStrikeThrough (true);
            r3.setFontSize (20);
            r3.setSubscript (VerticalAlign.SUPERSCRIPT);

            // гиперссылка
            XWPFHyperlinkRun hyperlink = p2.insertNewHyperlinkRun (0, "http://poi.apache.org/");
            hyperlink.setUnderline (UnderlinePatterns.SINGLE); 
            hyperlink.setColor ("0000FF"); 
            hyperlink.setText ("POI Apache"); 

            XWPFParagraph p3 = doc.createParagraph (); 
            p3.setWordWrapped (true); 
            p3.setPageBreak (true); 

            //p3.setAlignment(ParagraphAlignment.DISTRIBUTE); 
            p3.setAlignment (ParagraphAlignment.BOTH); 
            p3.setSpacingBetween(15, LineSpacingRule.EXACT);

            p3.setIndentationFirstLine (600); 


            XWPFRun r4 = p3.createRun (); 
            r4.setTextPosition (20);
            r4.setText ("To be or not to be: that is the question:" 
                    + "Be nobler in your mind to suffer"
                    + “Slings and arrows of mad luck”, 
                    + “Or take arms against the sea of ​​troubles”, 
                    + “And the end of those who oppose them? Die: sleep; "); 
            r4.addBreak (BreakType.PAGE); 
            r4.setText ("No more; and sleep to say that we are finished" 
                    + "Pain in the heart and a thousand natural upheavals" 
                    + "This flesh is the heir, this is completion" 
                    + "Truly desire. Die, sleep;" 
                    + " Sleep: perhaps dream: ay, that's the catch; " 
                    +" ....... "); 
            r4.setItalic (true);
// This would mean that this break should be treated as a simple line break and a line break after this word: 

            XWPFRun r5 = p3.createRun (); 
            r5.setTextPosition (-10); 
            r5.setText ("For there may be dreams in that death dream"); 
            r5.addCarriageReturn (); 
            r5.setText ("When we finish this deadly coil," 
                    + "We must pause: here comes the respect" 
                    + "This is the trouble of such a long life;"); 
            r5.addBreak (); 
            r5.setText ("For those who endure the whips and contempt of time," 
                    + "The oppressor is wrong, the proud is offended"); 

            r5.addBreak (BreakClear.ALL);
            r5.setText ("Torment of despicable love, delay with the law" 
                    + "Insolence of office and neglect" + "......."); 

            try (FileOutputStream out = new FileOutputStream ("simple.docx")) { 
                doc.write (out); 
            } 
        } 
    } 
}