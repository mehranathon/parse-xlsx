package com.poc.parse_xlsx.service;

import com.poc.parse_xlsx.model.Person;
import java.io.IOException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.List;
import java.util.stream.IntStream;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
@Slf4j
public class UploadService {
  @Value("${excel-template-name}")
  private String templateName;

  public List<Person> parseXlsx(MultipartFile file) {
    log.info("parsing: {}", file.getOriginalFilename());

    XSSFWorkbook workbook;
    try {
      workbook = new XSSFWorkbook(file.getInputStream());
    } catch (IOException e) {
      log.error(e.getMessage());
      return null;
    }

    XSSFTable table = workbook.getTable(templateName);
    XSSFSheet sheet = table.getXSSFSheet();

    int headerRow = table.getStartRowIndex();
    int finalRow = table.getEndRowIndex();
    int totalRows = finalRow - headerRow;

    // corrects for vertical but not horizontal offset
    return IntStream.range(0, totalRows)
        .mapToObj(i -> rowToModel(sheet.getRow(i + headerRow + 1)))
        .toList();
  }

  private Person rowToModel(Row row) {
    // assumes correct order of columns
    Person person = new Person();
    Cell idCell = row.getCell(3);

    person.setFirstname(row.getCell(0).toString());
    person.setLastName(row.getCell(1).toString());
    person.setDob(
        LocalDate.ofInstant(
                row.getCell(2).getDateCellValue().toInstant(),
                ZoneId.systemDefault()));
    // apache poi may pick up STRING as NUMERIC (double)
    person.setIdentifier(
        idCell.getCellType() == CellType.NUMERIC
            ? String.valueOf((int) idCell.getNumericCellValue())
            : idCell.toString());

    return person;
  }
}
