package com.example.studentbackend.service;

import com.example.studentbackend.entity.Student;
import com.example.studentbackend.repository.StudentRepository;
import com.opencsv.CSVWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@Service
public class CSVService {

    private final StudentRepository repo;

    public CSVService(StudentRepository repo) {
        this.repo = repo;
    }

    public File convertExcelToCSV(MultipartFile excelFile, String csvPath) throws Exception {
        Workbook workbook = new XSSFWorkbook(excelFile.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        File csvFile = new File(csvPath);
        csvFile.getParentFile().mkdirs();
        CSVWriter writer = new CSVWriter(new FileWriter(csvFile));

        for (Row row : sheet) {
            List<String> data = new ArrayList<>();
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.NUMERIC) data.add(String.valueOf((int) cell.getNumericCellValue() + (row.getRowNum() == 0 ? 0 : 10)));
                else data.add(cell.getStringCellValue());
            }
            writer.writeNext(data.toArray(new String[0]));
        }

        writer.close();
        workbook.close();
        return csvFile;
    }

    public void saveCSVToDB(MultipartFile csvFile) throws Exception {
        BufferedReader reader = new BufferedReader(new InputStreamReader(csvFile.getInputStream()));
        String line;
        boolean header = true;
        while ((line = reader.readLine()) != null) {
            if (header) { header = false; continue; }
            String[] cols = line.split(",");
            Student s = new Student();
            s.setFirstName(cols[1]);
            s.setLastName(cols[2]);
            s.setDob(LocalDate.parse(cols[3]));
            s.setStudentClass(cols[4]);
            s.setScore(Integer.parseInt(cols[5]) + 5);
            repo.save(s);
        }
        reader.close();
    }
}
