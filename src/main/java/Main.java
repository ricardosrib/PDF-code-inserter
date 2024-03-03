import org.apache.pdfbox.pdmodel.font.PDType1Font;

import java.io.IOException;
import java.util.List;

public class Main {

    public static void main(String[] args) {

        int sheetIndex = 0;
        int columnIndex = 0;
        int rowIndex = 0;

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
                    sheetIndex,
                    columnIndex,
                    rowIndex
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
