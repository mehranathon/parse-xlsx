package com.poc.parse_xlsx.controller;

import com.poc.parse_xlsx.model.Person;
import com.poc.parse_xlsx.service.UploadService;
import java.util.List;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@Slf4j
public class UploadController {
  @Autowired UploadService uploadService;

  @PostMapping(value = "/upload_xslx", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
  public List<Person> uploadXlsx(@RequestPart("file") MultipartFile file) {
    return uploadService.parseXlsx(file);
  }
}
