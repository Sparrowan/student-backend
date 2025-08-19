package com.example.studentbackend.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.opencsv.CSVWriter;
import org.springframework.stereotype.Service;

import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;

@Service
public class ExcelToCSVService {

    public void convertExcelToCSV(String excelFile, String csvFile) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);
        try (CSVWriter writer = new CSVWriter(new FileWriter(csvFile))) {
            for (Row row : sheet) {
                String[] data = new String[row.getPhysicalNumberOfCells()];
                for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
                    Cell cell = row.getCell(i);
                    data[i] = cell.toString();
                }
                writer.writeNext(data);
            }
        }
        workbook.close();
    }
}
