package com.sas.springbootexceldownload.entity;

import lombok.*;

@Getter
@Setter
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class Employee {

    private Long employeeId;
    private String firstName;
    private String lastName;
    private String designation;
    private String department;

}
