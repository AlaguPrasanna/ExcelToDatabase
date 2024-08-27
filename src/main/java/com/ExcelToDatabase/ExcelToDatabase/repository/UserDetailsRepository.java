package com.ExcelToDatabase.ExcelToDatabase.repository;

import com.ExcelToDatabase.ExcelToDatabase.model.UserDetails;
import org.springframework.data.jpa.repository.JpaRepository;

public interface UserDetailsRepository extends JpaRepository<UserDetails,Long> {

}
