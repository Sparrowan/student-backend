package com.example.studentbackend.service;

import com.example.studentbackend.entity.Student;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class ExcelService {

    private SXSSFWorkbook workbook;
    private Sheet sheet;
    private int currentRow = 0;

    // Initialize streaming Excel
    public File createStreamingExcel(String path) throws IOException {
        workbook = new SXSSFWorkbook(100); // keep 100 rows in memory
        sheet = workbook.createSheet("Students");

        // Create header row
        Row header = sheet.createRow(currentRow++);
        header.createCell(0).setCellValue("ID");
        header.createCell(1).setCellValue("First Name");
        header.createCell(2).setCellValue("Last Name");
        header.createCell(3).setCellValue("Class");
        header.createCell(4).setCellValue("Score");
        header.createCell(5).setCellValue("DOB");

        return new File(path);
    }

    // Write a single student row
    public void writeStudentRow(File file, Student s, int id) throws IOException {
        Row row = sheet.createRow(currentRow++);
        row.createCell(0).setCellValue(id);
        row.createCell(1).setCellValue(s.getFirstName());
        row.createCell(2).setCellValue(s.getLastName());
        row.createCell(3).setCellValue(s.getStudentClass());
        row.createCell(4).setCellValue(s.getScore());
        row.createCell(5).setCellValue(s.getDob().toString());
    }

    // Flush and close the workbook
    public void finishStreamingExcel(File file) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        } finally {
            workbook.dispose(); // remove temporary files
            workbook.close();
        }
    }
}
