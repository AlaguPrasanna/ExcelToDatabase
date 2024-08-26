package com.ExcelToDatabase.ExcelToDatabase.util;

import com.ExcelToDatabase.ExcelToDatabase.model.Users;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelHelper {
    private static final String SHEET = "Users";
    private static final String[] HEADERs = { "UserName", "EmailId" };
    public static final String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


    public static List<Users> excelToUsers(InputStream is) {
        try (XSSFWorkbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            if (sheet == null) {
                throw new RuntimeException("Sheet is null in the provided Excel file.");
            }

            List<Users> users = new ArrayList<>();
            Iterator<Row> rows = sheet.iterator();

            // Skip header row
            if (rows.hasNext()) {
                rows.next();
            }

            while (rows.hasNext()) {
                Row row = rows.next();
                Users user = new Users();

                Cell userNameCell = row.getCell(0);
                Cell emailIdCell = row.getCell(1);

                if (userNameCell != null) {
                    user.setUserName(userNameCell.getStringCellValue());
                }
                if (emailIdCell != null) {
                    user.setEmailId(emailIdCell.getStringCellValue());
                }

                users.add(user);
            }
            return users;
        } catch (IOException e) {
            throw new RuntimeException("Failed to parse Excel file: " + e.getMessage(), e);
        }
    }

    public static ByteArrayInputStream usersToExcel(List<Users> users) {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet(SHEET);

            // Header
            Row headerRow = sheet.createRow(0);
            for (int col = 0; col < HEADERs.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(HEADERs[col]);
            }

            // Data rows
            int rowIdx = 1;
            for (Users user : users) {
                Row row = sheet.createRow(rowIdx++);

                row.createCell(0).setCellValue(user.getUserName());
                row.createCell(1).setCellValue(user.getEmailId());
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("Failed to write Excel file: " + e.getMessage());
        }
}}


