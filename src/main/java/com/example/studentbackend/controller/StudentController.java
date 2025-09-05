package com.example.studentbackend.controller;

import com.example.studentbackend.entity.Student;
import com.example.studentbackend.repository.StudentRepository;
import com.example.studentbackend.service.CSVService;
import com.example.studentbackend.service.ExcelService;
import com.example.studentbackend.service.ReportService;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;

import java.io.File;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/students")
public class StudentController {

    private final ExcelService excelService;
    private final StudentRepository studentRepository;

    @Autowired
    private CSVService csvService;

    @Autowired
    private ReportService reportService;

    public StudentController(ExcelService excelService, StudentRepository studentRepository) {
        this.excelService = excelService;
        this.studentRepository = studentRepository;
    }

    // Generate test data + Excel file
    @PostMapping("/generate-excel")
    public String generateExcel(@RequestParam int count) throws Exception {
        int batchSize = 10000;
        List<Student> batch = new ArrayList<>(batchSize);

        File file = excelService.createStreamingExcel("data/students.xlsx");

        for (int i = 1; i <= count; i++) {
            Student s = new Student();
            s.setFirstName(randomString(3, 8));
            s.setLastName(randomString(3, 8));
            s.setDob(randomDate());
            s.setStudentClass("Class" + ((i % 5) + 1));
            s.setScore(55 + (int) (Math.random() * 21));

            batch.add(s);
            excelService.writeStudentRow(file, s, i);

            if (batch.size() >= batchSize) {
                studentRepository.saveAll(batch);
                batch.clear();
            }
        }

        if (!batch.isEmpty()) {
            studentRepository.saveAll(batch);
        }

        excelService.finishStreamingExcel(file);

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
        List<Student> students = studentRepository.findAll();
        reportService.exportExcel(students, "data/report.xlsx");
        return "Report Excel generated";
    }

    @GetMapping("/report/csv")
    public String exportCSV() throws Exception {
        List<Student> students = studentRepository.findAll();
        reportService.exportCSV(students, "data/report.csv");
        return "Report CSV generated";
    }

    @GetMapping("/report/pdf")
    public String exportPDF() throws Exception {
        List<Student> students = studentRepository.findAll();
        reportService.exportPDF(students, "data/report.pdf");
        return "Report PDF generated";
    }

    // List all students with pagination
    @GetMapping
    public Page<Student> getAllStudents(Pageable pageable) {
        return studentRepository.findAll(pageable);
    }

    // Search by ID
    @GetMapping("/search/{id}")
    public Student getStudentById(@PathVariable Long id) {
        return studentRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Student not found with id " + id));
    }

    // Filter by class with pagination
    @GetMapping("/class/{studentClass}")
    public Page<Student> getStudentsByClass(@PathVariable String studentClass, Pageable pageable) {
        return studentRepository.findByStudentClass(studentClass, pageable);
    }

    

    // Utilities
    private String randomString(int minLen, int maxLen) {
        int len = minLen + (int) (Math.random() * (maxLen - minLen + 1));
        StringBuilder sb = new StringBuilder();
        String alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        for (int i = 0; i < len; i++) {
            sb.append(alphabet.charAt((int) (Math.random() * alphabet.length())));
        }
        return sb.toString();
    }

    private LocalDate randomDate() {
        int year = 2000 + (int) (Math.random() * 11);
        int month = 1 + (int) (Math.random() * 12);
        int day = 1 + (int) (Math.random() * 28);
        return LocalDate.of(year, month, day);
    }
}
