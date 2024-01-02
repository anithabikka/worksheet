package com.ms.worksheet.controller;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("/api/excel")
public class WorkSheetApi {

    @GetMapping("/generate")
    public ResponseEntity<byte[]> generateExcel() {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Employee Data");

            // Create headers (first row)
            Row headerRow = sheet.createRow(0);
            CellStyle centeredStyle2 = workbook.createCellStyle();
            centeredStyle2.setAlignment(HorizontalAlignment.CENTER);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 3));
            Cell one=headerRow.createCell(2);
            one.setCellValue("Header 1");
            one.setCellStyle(centeredStyle2);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 5));
            Cell two =headerRow.createCell(4);
            two.setCellValue("Header 2");
            two.setCellStyle(centeredStyle2);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 6, 7));
            Cell three =headerRow.createCell(6);
            three.setCellValue("Header 3");
            three.setCellStyle(centeredStyle2);
            Cell four =headerRow.createCell(8);
            four.setCellValue("Header 4");
            four.setCellStyle(centeredStyle2);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 9, 10));
            Cell five =headerRow.createCell(9);
            five.setCellValue("Header 5");
            five.setCellStyle(centeredStyle2);


            // Create data (second row)
            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue("Employee No");
            dataRow.createCell(1).setCellValue("Name");
            dataRow.createCell(2).setCellValue("SubHeader 1");
            dataRow.createCell(3).setCellValue("SubHeader 2");
            dataRow.createCell(4).setCellValue("SubHeader 1");
            dataRow.createCell(5).setCellValue("SubHeader 2");
            dataRow.createCell(6).setCellValue("SubHeader 1");
            dataRow.createCell(7).setCellValue("SubHeader 2");
            dataRow.createCell(8).setCellValue("SubHeader 3");
            dataRow.createCell(9).setCellValue("SubHeader 1");
            dataRow.createCell(10).setCellValue("SubHeader 2");



            byte[] excelBytes;
            try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
                workbook.write(byteArrayOutputStream);
                excelBytes = byteArrayOutputStream.toByteArray();
            }

            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "employee_data.xlsx");

            return new ResponseEntity<>(excelBytes, headers, HttpStatus.OK);
        } catch (IOException e) {
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}
