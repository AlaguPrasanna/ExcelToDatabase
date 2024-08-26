package com.ExcelToDatabase.ExcelToDatabase.service;

import com.ExcelToDatabase.ExcelToDatabase.model.Users;
import com.ExcelToDatabase.ExcelToDatabase.repository.UserRepository;
import com.ExcelToDatabase.ExcelToDatabase.util.ExcelHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

@Service
public class UserService {

    @Autowired
    private UserRepository userRepository;


    public void saveUsersFromExcel(MultipartFile file) {
        try {
            List<Users> users = ExcelHelper.excelToUsers(file.getInputStream());
            userRepository.saveAll(users);
        } catch (IOException e) {
            throw new RuntimeException("Failed to store Excel data: " + e.getMessage());
        }
    }

    public ByteArrayInputStream loadUsersToExcel() {
        List<Users> users = userRepository.findAll();
        return ExcelHelper.usersToExcel(users);
    }

}