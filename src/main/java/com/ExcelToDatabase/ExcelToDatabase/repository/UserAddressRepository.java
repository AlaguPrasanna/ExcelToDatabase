package com.ExcelToDatabase.ExcelToDatabase.repository;

import com.ExcelToDatabase.ExcelToDatabase.model.UserAddress;
import org.springframework.data.jpa.repository.JpaRepository;

public interface UserAddressRepository extends JpaRepository<UserAddress,Long> {
}
