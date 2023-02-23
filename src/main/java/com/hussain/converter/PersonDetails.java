package com.hussain.converter;

import com.hussain.response.Person;

import java.util.List;

public class PersonDetails {

    public static List<Person> getPersons() {
        return List.of(new Person("XYZ", "xyz@gmail.com", "Kolkata"),
                new Person("XYZ1", "xyz1@gmail.com", "Mumbai"),
                new Person("XYZ2", "xyz2@gmail.com", "Delhi"),
                new Person("XYZ3", "xyz3@gmail.com", "Chennai"),
                new Person("XYZ4", "xyz4e@gmail.com", "Bangalore"),
                new Person("XYZ5", "XYZ5@gmail.com", "Hyderabad"));
    }
}
