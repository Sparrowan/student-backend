package com.example.studentbackend.service;

import com.example.studentbackend.entity.Student;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.List;

@Service
public class ExcelService {

    public File generateExcel(List<Student> students, String filePath) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Students");

        Row header = sheet.createRow(0);
        String[] columns = {"ID", "First Name", "Last Name", "DOB", "Class", "Score"};
        for (int i = 0; i < columns.length; i++) header.createCell(i).setCellValue(columns[i]);

        int rowNum = 1;
        for (Student s : students) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(s.getId());
            row.createCell(1).setCellValue(s.getFirstName());
            row.createCell(2).setCellValue(s.getLastName());
            row.createCell(3).setCellValue(s.getDob().toString());
            row.createCell(4).setCellValue(s.getStudentClass());
            row.createCell(5).setCellValue(s.getScore());
        }

        File file = new File(filePath);
        file.getParentFile().mkdirs();
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        workbook.close();
        fos.close();

        return file;
    }
}
