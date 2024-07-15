package serviceImpl;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PdfToExcelConverter {
    public static void main(String[] args) {
        try {
            // Load PDF document
            PDDocument document = PDDocument.load(new File("C:\\DataFiles\\Cibil Co-applicant 1.pdf"));

            // Create PDFTextStripper
            PDFTextStripper pdfStripper = new PDFTextStripper();

            // Extract text from PDF
            String pdfText = pdfStripper.getText(document);

            // Create Excel workbook
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Data");

            // Write extracted data to Excel
            String[] lines = pdfText.split("\\r?\\n");
            int rowNum = 0;
            for (String line : lines) {
                Row row = sheet.createRow(rowNum++);
                String[] columns = line.split(",");
                int colNum = 0;
                for (String column : columns) {
                    row.createCell(colNum++).setCellValue(column);
                }
            }

            // Write Excel workbook to file
            FileOutputStream outputStream = new FileOutputStream(new File("C:\\DataFiles\\output.xlsx"));
            workbook.write(outputStream);
            workbook.close();

            // Close PDF document
            document.close();

            System.out.println("Data extracted from PDF and written to Excel successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
