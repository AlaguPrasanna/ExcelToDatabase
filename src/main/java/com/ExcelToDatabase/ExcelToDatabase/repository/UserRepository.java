package com.ExcelToDatabase.ExcelToDatabase.repository;

import com.ExcelToDatabase.ExcelToDatabase.model.Users;
import org.springframework.data.jpa.repository.JpaRepository;

public interface UserRepository extends JpaRepository<Users, Long> {
}
