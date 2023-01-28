package com.sas.springbootexceldownload.entity;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class Contact {
    private String firstName;
    private String lastName;
    private String email;
    private String phoneNumber;
    private String address;
}
