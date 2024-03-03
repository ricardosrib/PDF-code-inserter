import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.IOException;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        String excelFilePath = "";
        String pdfFilePath = "";
        int pageNum = 1;
        int xPosition = 410;
        int yPosition = 743;
        PDType1Font fontType = PDType1Font.HELVETICA_BOLD;
        int fontSize = 12;

        try {
            // Read codes from Excel file
            List<String> codes = PDFCodeInserter.readCodesFromExcel(
                    excelFilePath,
                    0,
                    0,
                    0
            );

            // Process each PDF file
            for (String code : codes) {
                PDFCodeInserter.insertCodeIntoPDF(
                        pdfFilePath,
                        code,
                        "" + code + ".pdf",
                        pageNum,
                        xPosition,
                        yPosition,
                        fontType,
                        fontSize
                );
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
