import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import java.io.File;


public class PDFCodeInserter {

    static List<String> readCodesFromExcel(
            String excelFilePath,
            int sheetIndex,
            int columnIndex,
            int rowIndex
    ) {
        List<String> codes = new ArrayList<>();

        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> rowIterator = sheet.iterator();

            for (int i = 0; i < rowIndex; i++) {
                rowIterator.next();
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(columnIndex);

                if (cell != null) {
                    // Assuming the codes are numeric (adjust if they are of different types)
                    codes.add(String.valueOf((long) cell.getNumericCellValue()));
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return codes;
    }

    static void insertCodeIntoPDF(
            String pdfFilePath,
            String code,
            String outputFilePath,
            int pageNum,
            float x,
            float y,
            PDType1Font fontType,
            int fontSize
    ) throws IOException {

        try (PDDocument document = PDDocument.load(new File(pdfFilePath))) {
            PDPage page = document.getPage(pageNum - 1);
            try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true)) {
                contentStream.beginText();
                contentStream.setFont(fontType, fontSize);
                contentStream.newLineAtOffset(x, y);
                contentStream.showText(code);
                contentStream.endText();
            }

            document.save(outputFilePath);
        }
    }
}
