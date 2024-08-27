package com.ExcelToDatabase.ExcelToDatabase.util;

import com.ExcelToDatabase.ExcelToDatabase.model.UserAddress;
import com.ExcelToDatabase.ExcelToDatabase.model.UserDetails;
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
    private static final String[] HEADERs = { "Username", "Email", "Phone Number", "Street", "City", "State", "Zip Code" };
    private static final String SHEET = "Users";

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
                UserDetails userDetails = new UserDetails();
                UserAddress userAddress = new UserAddress();

                Cell userNameCell = row.getCell(0);
                Cell emailIdCell = row.getCell(1);
                Cell phoneNumberCell = row.getCell(2);
                Cell streetCell = row.getCell(3);
                Cell cityCell = row.getCell(4);
                Cell stateCell = row.getCell(5);
                Cell zipCodeCell = row.getCell(6);

                if (userNameCell != null) {
                    user.setUserName(getCellValue(userNameCell));
                }
                if (emailIdCell != null) {
                    user.setEmailId(getCellValue(emailIdCell));
                }
                if (phoneNumberCell != null) {
                    userDetails.setPhoneNumber(getCellValue(phoneNumberCell));
                }
                if (streetCell != null) {
                    userAddress.setStreet(getCellValue(streetCell));
                }
                if (cityCell != null) {
                    userAddress.setCity(getCellValue(cityCell));
                }
                if (stateCell != null) {
                    userAddress.setState(getCellValue(stateCell));
                }
                if (zipCodeCell != null) {
                    userAddress.setZipCode(getCellValue(zipCodeCell));
                }

                user.setUserDetails(userDetails);
                user.setUserAddress(userAddress);

                users.add(user);
            }
            return users;
        } catch (IOException e) {
            throw new RuntimeException("Failed to parse Excel file: " + e.getMessage(), e);
        }
    }

    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
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

                if (user.getUserDetails() != null) {
                    row.createCell(2).setCellValue(user.getUserDetails().getPhoneNumber());
                } else {
                    row.createCell(2).setCellValue("");
                }

                if (user.getUserAddress() != null) {
                    row.createCell(3).setCellValue(user.getUserAddress().getStreet());
                    row.createCell(4).setCellValue(user.getUserAddress().getCity());
                    row.createCell(5).setCellValue(user.getUserAddress().getState());
                    row.createCell(6).setCellValue(user.getUserAddress().getZipCode());
                } else {
                    row.createCell(3).setCellValue("");
                    row.createCell(4).setCellValue("");
                    row.createCell(5).setCellValue("");
                    row.createCell(6).setCellValue("");
                }
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("Failed to write Excel file: " + e.getMessage());
        }
    }
}


