package com.ExcelToDatabase.ExcelToDatabase.controller;


import com.ExcelToDatabase.ExcelToDatabase.service.UserService;
import com.ExcelToDatabase.ExcelToDatabase.util.ExcelHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import java.io.ByteArrayInputStream;

@RestController
@RequestMapping("/api/users")
public class UserController {

    @Autowired
    private UserService userService;

    @PostMapping("/upload/excel")
    public ResponseEntity<String> uploadExcel(@RequestParam("file") MultipartFile file) {
        userService.saveUsersFromExcel(file);
        return ResponseEntity.ok("Excel file uploaded and data saved successfully.");
    }

    @GetMapping("/download/excel")
    public ResponseEntity<?> downloadExcel() {
        ByteArrayInputStream in = userService.loadUsersToExcel();
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=users.xlsx")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(new org.springframework.core.io.InputStreamResource(in));
    }
}