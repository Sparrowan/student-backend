package com.example.studentbackend.controller;

import com.example.studentbackend.entity.Student;
import com.example.studentbackend.repository.StudentRepository;
import com.example.studentbackend.service.CSVService;
import com.example.studentbackend.service.ExcelService;
import com.example.studentbackend.service.ReportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/students")
public class StudentController {

    @Autowired
    private ExcelService excelService;

    @Autowired
    private CSVService csvService;

    @Autowired
    private ReportService reportService;

    @Autowired
    private StudentRepository repo;

    // Generate Excel with N random students
@PostMapping("/generate-excel")
public String generateExcel(@RequestParam int count) throws Exception {
    List<Student> students = new ArrayList<>();
    for (int i = 1; i <= count; i++) {
        Student s = new Student();
        s.setFirstName(randomString(3,8));
        s.setLastName(randomString(3,8));
        s.setDob(randomDate());
        s.setStudentClass("Class" + ((i % 5) + 1));
        s.setScore(55 + (int)(Math.random()*21));
        students.add(s);
    }

    // Save students to DB so they get IDs
    students = repo.saveAll(students);

    // Now generate Excel
    File file = excelService.generateExcel(students, "data/students.xlsx");
    return "Excel generated at: " + file.getAbsolutePath();
}


    // Convert Excel → CSV
    @PostMapping("/excel-to-csv")
    public String excelToCSV(@RequestParam MultipartFile file) throws Exception {
        File csv = csvService.convertExcelToCSV(file, "data/students.csv");
        return "CSV generated at: " + csv.getAbsolutePath();
    }

    // Upload CSV → DB
    @PostMapping("/upload-csv")
    public String uploadCSV(@RequestParam MultipartFile file) throws Exception {
        csvService.saveCSVToDB(file);
        return "CSV uploaded to DB";
    }

    // Export reports
    @GetMapping("/report/excel")
    public String exportExcel() throws Exception {
        List<Student> students = repo.findAll();
        reportService.exportExcel(students, "data/report.xlsx");
        return "Report Excel generated";
    }

    @GetMapping("/report/csv")
    public String exportCSV() throws Exception {
        List<Student> students = repo.findAll();
        reportService.exportCSV(students, "data/report.csv");
        return "Report CSV generated";
    }

    @GetMapping("/report/pdf")
    public String exportPDF() throws Exception {
        List<Student> students = repo.findAll();
        reportService.exportPDF(students, "data/report.pdf");
        return "Report PDF generated";
    }

    // Utilities
    private String randomString(int min, int max) {
        int len = min + (int)(Math.random()*(max-min+1));
        StringBuilder sb = new StringBuilder();
        String chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        for(int i=0;i<len;i++) sb.append(chars.charAt((int)(Math.random()*chars.length())));
        return sb.toString();
    }

    private LocalDate randomDate() {
        int startYear = 2000, endYear = 2010;
        int day = 1 + (int)(Math.random()*28);
        int month = 1 + (int)(Math.random()*12);
        int year = startYear + (int)(Math.random()*(endYear-startYear+1));
        return LocalDate.of(year, month, day);
    }
}
