package com.example.studentbackend.repository;

import com.example.studentbackend.entity.Student;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.jpa.repository.JpaRepository;

public interface StudentRepository extends JpaRepository<Student, Long> {

    // Paging by student class
    Page<Student> findByStudentClass(String studentClass, Pageable pageable);

    // Optional: if you really want to page by ID (rarely needed)
    Page<Student> findAllById(Long studentId, Pageable pageable);
}
