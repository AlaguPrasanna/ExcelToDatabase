package com.ExcelToDatabase.ExcelToDatabase.service;

import com.ExcelToDatabase.ExcelToDatabase.model.Users;
import com.ExcelToDatabase.ExcelToDatabase.repository.UserAddressRepository;
import com.ExcelToDatabase.ExcelToDatabase.repository.UserDetailsRepository;
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
        private UserRepository usersRepository;

        @Autowired
        private UserDetailsRepository userDetailsRepository;

        @Autowired
        private UserAddressRepository userAddressRepository;

        public void saveUsersFromExcel(MultipartFile file) {
            try {
                List<Users> users = ExcelHelper.excelToUsers(file.getInputStream());
                for (Users user : users) {
                    if (!usersRepository.existsByEmailId(user.getEmailId())) {
                        // Save user and associated details
                        Users savedUser = usersRepository.save(user);

                        if (user.getUserDetails() != null) {
                            user.getUserDetails().setUser(savedUser);
                            userDetailsRepository.save(user.getUserDetails());
                        }

                        if (user.getUserAddress() != null) {
                            user.getUserAddress().setUser(savedUser);
                            userAddressRepository.save(user.getUserAddress());
                        }
                    }
                }
            } catch (Exception e) {
                throw new RuntimeException("Failed to store excel data: " + e.getMessage());
            }
        }

        public ByteArrayInputStream loadUsersToExcel() {
            List<Users> users = usersRepository.findAll();
            return ExcelHelper.usersToExcel(users);
        }
    }

