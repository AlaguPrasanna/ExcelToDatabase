package com.ExcelToDatabase.ExcelToDatabase.service;

import com.ExcelToDatabase.ExcelToDatabase.model.Users;
import com.ExcelToDatabase.ExcelToDatabase.repository.UserAddressRepository;
import com.ExcelToDatabase.ExcelToDatabase.repository.UserDetailsRepository;
import com.ExcelToDatabase.ExcelToDatabase.repository.UserRepository;
import com.ExcelToDatabase.ExcelToDatabase.util.ExcelHelper;
import org.slf4j.LoggerFactory;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;


@Service

    public class UserService {

    private static final Logger logger = LoggerFactory.getLogger(UserService.class);
    @Autowired
        private UserRepository usersRepository;

        @Autowired
        private UserDetailsRepository userDetailsRepository;

        @Autowired
        private UserAddressRepository userAddressRepository;
    public void saveUsersFromExcel(MultipartFile file) {
        logger.info("Started processing the Excel file for user data upload.");

        try {
            List<Users> usersFromExcel = ExcelHelper.excelToUsers(file.getInputStream());
            for (Users user : usersFromExcel) {
                // Check if user with the same email already exists
                if (!usersRepository.existsByEmailId(user.getEmailId())) {
                    logger.info("Saving user with email: {}", user.getEmailId());

                    // Save user, user details, and user address
                    usersRepository.save(user);
                    userDetailsRepository.save(user.getUserDetails());
                    userAddressRepository.save(user.getUserAddress());

                } else {
                    logger.warn("Duplicate user detected, skipping user with email: {}", user.getEmailId());
                }
            }
            logger.info("Finished processing the Excel file.");
        } catch (IOException e) {
            logger.error("Failed to store excel data due to an IO exception: {}", e.getMessage());
            throw new RuntimeException("Failed to store excel data: " + e.getMessage());
        } catch (Exception e) {
            logger.error("An unexpected error occurred: {}", e.getMessage());
            throw new RuntimeException("Failed to store excel data: " + e.getMessage());
        }
    }

    public ByteArrayInputStream loadUsersToExcel() {
        logger.info("Started exporting user data to Excel.");

        List<Users> users = usersRepository.findAll();
        ByteArrayInputStream excelData = ExcelHelper.usersToExcel(users);

        logger.info("Successfully exported user data to Excel.");
        return excelData;
    }

    }

