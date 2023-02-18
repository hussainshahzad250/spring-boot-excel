package com.hussain.service;

import com.hussain.response.Contact;
import com.hussain.response.Employee;
import com.hussain.utils.ContactDetails;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.List;

@Service
public class ExcelService {

    public boolean addSheet(String filePath, String newSheetName) {
        InputStream inputStream = null;
        Workbook workbook = null;
        OutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(filePath);
            workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.createSheet(newSheetName);
            Row row = sheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("Test New Cell");
            outputStream = new FileOutputStream(filePath);
            workbook.write(outputStream);
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return false;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        } finally {
            try {
                inputStream.close();
                workbook.close();
                outputStream.close();
            } catch (Exception ex) {
                ex.printStackTrace();
                return false;
            }
        }
    }

    public ByteArrayInputStream addNewSheet(MultipartFile multipartFile, String newSheetName) {
        InputStream inputStream = null;
        Workbook workbook = null;
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            inputStream = multipartFile.getInputStream();
            workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.createSheet(newSheetName);
            List<Contact> contacts = ContactDetails.getContacts();
            addContactDetails(workbook, sheet, contacts);
            workbook.write(outputStream);
            return new ByteArrayInputStream(outputStream.toByteArray());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
                workbook.close();
                outputStream.close();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return null;
    }

    private void addContactDetails(Workbook workbook, Sheet sheet, List<Contact> contacts) {

        Row row = sheet.createRow(0);

        // Define header cell style
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Creating header cells
        Cell cell = row.createCell(0);
        cell.setCellValue("First Name");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(1);
        cell.setCellValue("Last Name");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Email");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(3);
        cell.setCellValue("Phone Number");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(4);
        cell.setCellValue("Address");
        cell.setCellStyle(headerCellStyle);

        // Creating data rows for each contact
        for (int i = 0; i < contacts.size(); i++) {
            Row dataRow = sheet.createRow(i + 1);
            dataRow.createCell(0).setCellValue(contacts.get(i).getFirstName());
            dataRow.createCell(1).setCellValue(contacts.get(i).getLastName());
            dataRow.createCell(2).setCellValue(contacts.get(i).getEmail());
            dataRow.createCell(3).setCellValue(contacts.get(i).getPhoneNumber());
            dataRow.createCell(4).setCellValue(contacts.get(i).getAddress());
        }

        // Making size of column auto resize to fit with data
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
    }

    public ByteArrayInputStream generateExcel(List<Contact> contacts,List<Employee> employees) {
        Workbook workbook = null;
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            workbook = new XSSFWorkbook();
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Sheet sheet = workbook.createSheet("Contact Details");
            addContactDetails(workbook, sheet, contacts);
            sheet = workbook.createSheet("Employee Details");
            addEmployeeDetails(workbook, sheet, employees);
            workbook.write(outputStream);
            return new ByteArrayInputStream(outputStream.toByteArray());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
                outputStream.close();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return null;
    }

    private void addEmployeeDetails(Workbook workbook, Sheet sheet, List<Employee> employeeDetails) {
        Row row = sheet.createRow(0);

        // Define header cell style
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Creating header cells
        Cell cell = row.createCell(0);
        cell.setCellValue("Emp Id");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(1);
        cell.setCellValue("First Name");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Last Name");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(3);
        cell.setCellValue("Designation");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(4);
        cell.setCellValue("Department");
        cell.setCellStyle(headerCellStyle);

        // Creating data rows for each contact
        for (int i = 0; i < employeeDetails.size(); i++) {
            Row dataRow = sheet.createRow(i + 1);
            dataRow.createCell(0).setCellValue(employeeDetails.get(i).getEmployeeId());
            dataRow.createCell(1).setCellValue(employeeDetails.get(i).getFirstName());
            dataRow.createCell(2).setCellValue(employeeDetails.get(i).getLastName());
            dataRow.createCell(3).setCellValue(employeeDetails.get(i).getDesignation());
            dataRow.createCell(4).setCellValue(employeeDetails.get(i).getDepartment());
        }

        // Making size of column auto resize to fit with data
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
    }


}
