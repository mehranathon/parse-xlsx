package com.poc.parse_xlsx.model;

import java.time.LocalDate;
import lombok.Data;

@Data
public class Person {
  String firstname;
  String lastName;
  LocalDate dob;
  String identifier;
}
