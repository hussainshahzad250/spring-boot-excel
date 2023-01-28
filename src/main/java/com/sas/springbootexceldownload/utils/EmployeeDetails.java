package com.sas.springbootexceldownload.utils;


import com.sas.springbootexceldownload.entity.Employee;

import java.util.List;

public class EmployeeDetails {
    public static List<Employee> getEmployeeDetails() {
        return List.of(new Employee(1000L,"Shahzad","Hussain","",""));
    }
}
