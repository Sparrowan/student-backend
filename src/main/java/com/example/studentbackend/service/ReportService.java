package com.example.studentbackend.service;

import com.example.studentbackend.entity.Student;
import com.lowagie.text.Document;
import com.lowagie.text.Element;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

@Service
public class ReportService {

    // ================= Excel Export =================
    public void exportExcel(List<Student> students, String excelFilePath) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Students");

        int rowNum = 0;
        // Header
        Row headerRow = sheet.createRow(rowNum++);
        String[] headers = {"ID", "First Name", "Last Name", "Class", "Score"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Data rows
        for (Student s : students) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(s.getId());
            row.createCell(1).setCellValue(s.getFirstName());
            row.createCell(2).setCellValue(s.getLastName());
            row.createCell(4).setCellValue(s.getStudentClass());
            row.createCell(5).setCellValue(s.getScore());
        }

        // Write to file
        try (FileOutputStream fos = new FileOutputStream(new File(excelFilePath))) {
            workbook.write(fos);
        }
        workbook.close();
    }

    // ================= CSV Export =================
    public void exportCSV(List<Student> students, String csvFilePath) throws Exception {
        StringBuilder sb = new StringBuilder();
        sb.append("ID,First Name,Last Name,Class,Score\n");

        for (Student s : students) {
            sb.append(s.getId()).append(",");
            sb.append(s.getFirstName()).append(",");
            sb.append(s.getLastName()).append(",");
            sb.append(s.getStudentClass()).append(",");
            sb.append(s.getScore()).append("\n");
        }

        try (FileOutputStream fos = new FileOutputStream(new File(csvFilePath))) {
            fos.write(sb.toString().getBytes());
        }
    }

    // ================= PDF Export =================
    public void exportPDF(List<Student> students, String pdfFilePath) throws Exception {
        Document document = new Document();
        PdfWriter.getInstance(document, new FileOutputStream(new File(pdfFilePath)));
        document.open();

        Paragraph title = new Paragraph("Student Report");
        title.setAlignment(Element.ALIGN_CENTER);
        document.add(title);
        document.add(new Paragraph(" "));

        PdfPTable table = new PdfPTable(6); // 6 columns
        table.addCell("ID");
        table.addCell("First Name");
        table.addCell("Last Name");
        table.addCell("Class");
        table.addCell("Score");

        for (Student s : students) {
            table.addCell(String.valueOf(s.getId()));
            table.addCell(s.getFirstName());
            table.addCell(s.getLastName());
            table.addCell(s.getStudentClass());
            table.addCell(String.valueOf(s.getScore()));
        }

        document.add(table);
        document.close();
    }
}
