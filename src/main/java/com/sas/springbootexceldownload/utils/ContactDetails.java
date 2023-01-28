package com.sas.springbootexceldownload.utils;

import com.sas.springbootexceldownload.entity.Contact;

import java.util.List;

public class ContactDetails {


    public static List<Contact> getContacts() {
        return List.of(
                new Contact("Shahzad", "Hussain", "abc@gmail.com", "1234567890", "Gurgaon"),
                new Contact("Ejaz", "Hussain", "abc@gmail.com", "1234567890", "Gurgaon"),
                new Contact("Sail", "Hussain", "abc@gmail.com", "1234567890", "Gurgaon")
        );
    }
}
