
package sendmai;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class SendMails {

    // Regex pattern for email IDs
    private static final String EMAIL_REGEX = 
            "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,}";

    public static void main(String[] args) {
        // Change this folder path
        String folderPath = "/Users/vamsi/Documents/SendMails/SendMails/src/test/resources/InputFolder";
        String excelPath = "/Users/vamsi/Documents/SendMails/SendMails/src/test/resources/OutPutFolder/Output.xlsx";

        List<String> allEmails = new ArrayList<>();

        File folder = new File(folderPath);
        File[] listOfFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));

        if (listOfFiles == null || listOfFiles.length == 0) {
            System.out.println("No PDF files found in folder!");
            return;
        }

        for (File file : listOfFiles) {
            System.out.println("Reading: " + file.getName());
            try {
                List<String> emails = extractEmailsFromPDF(file);
                allEmails.addAll(emails);
            } catch (IOException e) {
                System.err.println("Error reading file " + file.getName() + ": " + e.getMessage());
            }
        }

        // Remove duplicates
        Set<String> uniqueEmails = new LinkedHashSet<>(allEmails);

        // Write to Excel
        try {
            writeEmailsToExcel(uniqueEmails, excelPath);
            System.out.println("Emails successfully written to: " + excelPath);
        } catch (IOException e) {
            System.err.println("Error writing Excel file: " + e.getMessage());
        }
    }

    // Extract emails from a single PDF file
    private static List<String> extractEmailsFromPDF(File file) throws IOException {
        List<String> emails = new ArrayList<>();

        try (PDDocument document = PDDocument.load(file)) {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);

            Pattern pattern = Pattern.compile(EMAIL_REGEX);
            Matcher matcher = pattern.matcher(text);

            while (matcher.find()) {
                emails.add(matcher.group());
            }
        }

        return emails;
    }

    // Write email list into Excel
    private static void writeEmailsToExcel(Set<String> emails, String excelPath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Emails");

        int rowCount = 0;
        for (String email : emails) {
            Row row = sheet.createRow(rowCount++);
            Cell cell = row.createCell(0);
            cell.setCellValue(email);
        }

        // Auto-size column
        sheet.autoSizeColumn(0);

        try (FileOutputStream fos = new FileOutputStream(excelPath)) {
            workbook.write(fos);
        }
        workbook.close();
    }
}
