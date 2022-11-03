import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class readexcel2 extends Exception {
    static XWPFHyperlinkRun createHyperlinkRun(XWPFParagraph paragraph, String uri) {
        String rId = paragraph.getDocument().getPackagePart().addExternalRelationship(
                uri,
                XWPFRelation.HYPERLINK.getRelation()
        ).getId();
        CTHyperlink cthyperLink=paragraph.getCTP().addNewHyperlink();
        cthyperLink.setId(rId);
        cthyperLink.addNewR();
        return new XWPFHyperlinkRun(
                cthyperLink,
                cthyperLink.getRArray(0),
                paragraph
        );
    }
    public static void main(String[] args)
    {
        // Document doc = new Document();
        try
        {
//PDF
            //PdfWriter.getInstance(doc, new FileOutputStream("/Users/himanshig/Excel/Sep22.pdf"));
            // doc.open();
//Word
            XWPFDocument document = new XWPFDocument();
            FileOutputStream out = new FileOutputStream("/u01/jenkins/workspace/release_announcement_document/Release_Announcement.docx");

//header of file

            FileReader reader=new FileReader("/Users/himanshig/Excel/untitled/src/db.properties");
            Properties p=new Properties();
            p.load(reader);
            XWPFParagraph paragraph9 = document.createParagraph();
            XWPFRun run = paragraph9.createRun();
            run.setText(p.getProperty("String1"));
            run.addBreak();
            paragraph9 = document.createParagraph();
            run = paragraph9.createRun();
            run.setText(p.getProperty("String2"));
            run.setBold(true);
            run.addBreak();
            paragraph9 = document.createParagraph();
            run = paragraph9.createRun();
            run.setText(p.getProperty("String3"));
            run.addBreak();

//Hashmap
            Map<String, Integer> requiredHeaders = new HashMap<>();
            FileInputStream file = new FileInputStream(new File("/u01/jenkins/workspace/release_announcement_document/Release_Announcement.xlsx"));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            DataFormatter formatter = new DataFormatter();
            Sheet sheet = workbook.getSheetAt(0);

            for (Cell cell : sheet.getRow(0)) {
                requiredHeaders.put(cell.getStringCellValue(), cell.getColumnIndex());
            }

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                String packageName=formatter.formatCellValue(row.getCell(requiredHeaders.get("Release Name")));
                String version = packageName.replaceAll(".*?((?<!\\w)\\d+([.-]\\d+)*).*", "$1");
//System.out.println(version);
                String ProductLine =formatter.formatCellValue(row.getCell(requiredHeaders.get("Product Line/Product")))+" "+ version;
                String CommunityURL=formatter.formatCellValue(row.getCell(requiredHeaders.get("Community URL")));

/*PDF WRITE
doc.add(new Paragraph("Product Line/Product = "+ProductLine));
doc.add(new Paragraph("Community URL = "+CommunityURL));

                XWPFParagraph paragraph1 = document.createParagraph();
                XWPFRun run = paragraph1.createRun();
                run.setText(ProductLine);*/
//Word Write
                //XWPFDocument document1 = new XWPFDocument();
                XWPFParagraph paragraph1 = document.createParagraph();
                XWPFRun run1 = paragraph1.createRun();
                run1.setText("");
                XWPFHyperlinkRun hyperlinkrun = createHyperlinkRun(paragraph1, CommunityURL);
                hyperlinkrun.setText(ProductLine);
                hyperlinkrun.setColor("0000FF");
                hyperlinkrun.setUnderline(UnderlinePatterns.SINGLE);
                //run = paragraph1.createRun();
                // run.setText("");
                // paragraph1 = document.createParagraph();
                //  paragraph1 = document.createParagraph();
                //  run = paragraph1.createRun();

// HYPERLINK of pdf
                //Paragraph paragraph = new Paragraph();
                //Anchor anchor = new Anchor(String.valueOf(doc.add(new Paragraph(ProductLine))));
                //Anchor anchor = new Anchor(ProductLine);
                //anchor.setReference(CommunityURL);
                //paragraph.add(anchor);
                //doc.add(paragraph);
//Print at console
                System.out.print(i+" "+ProductLine);
                System.out.print("   ||   ");
                System.out.println(CommunityURL);
                // System.out.print("Product Line/Product = " + formatter.formatCellValue(row.getCell(requiredHeaders.get("Product Line/Product")))+" "+ version);
                // System.out.print("   ||   ");
                // System.out.println("Community URL = " + formatter.formatCellValue(row.getCell(requiredHeaders.get("Community URL"))));

            }
//footer
            FileReader reader1=new FileReader("/Users/himanshig/Excel/untitled/src/db.properties");
            Properties p1=new Properties();
            p1.load(reader);
            XWPFParagraph paragraph10 = document.createParagraph();
            XWPFRun run10 = paragraph10.createRun();
            run10.addBreak();
            run10.setText(p.getProperty("String4"));
            run10.addBreak();
            paragraph10 = document.createParagraph();
            run10 = paragraph10.createRun();
            run10.setText(p.getProperty("String5"));

            // doc.close();
            workbook.close();
            document.write(out);
            out.close();
            System.out.println(" ");
            //System.out.println("Written to a pdf");
        }

        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}

