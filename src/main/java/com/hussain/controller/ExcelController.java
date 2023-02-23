package com.hussain.controller;

import com.hussain.service.ExcelService;
import com.hussain.converter.ContactDetails;
import com.hussain.converter.EmployeeDetails;
import org.apache.poi.util.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.IOException;

@RestController
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/addSheet")
    public String addSheet() {
        if (excelService.addSheet("D:\\QR-Code\\test.xlsx", "Test-Sheet"))
            return "Success";
        return "Sheet not added";
    }

    @PostMapping("/addSheetToExistingExcel")
    public void addSheetToExistingExcel(@RequestParam MultipartFile multipartFile, HttpServletResponse response) throws IOException {
        ByteArrayInputStream byteArrayInputStream = excelService.addNewSheet(multipartFile, "New Sheet");
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=contacts.xlsx");
        IOUtils.copy(byteArrayInputStream, response.getOutputStream());
    }

    @GetMapping("/createExcelWithMultipleSheet")
    public void createExcelWithMultipleSheet(HttpServletResponse response) throws IOException {
        ByteArrayInputStream byteArrayInputStream = excelService.generateExcel(ContactDetails.getContacts(), EmployeeDetails.getEmployeeDetails());
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=data.xlsx");
        IOUtils.copy(byteArrayInputStream, response.getOutputStream());
    }
}
