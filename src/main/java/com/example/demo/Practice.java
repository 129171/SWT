package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Practice {
    public static void main(String[] args) {
        String excelFilePath = "D:\\SWT\\lab1\\Lab1.xlsx";
        String sheetName = "Sheet1"; // Tên sheet trong file Excel

        try (FileInputStream inputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheet(sheetName);

            for (int i = 1; i <= 10000; i++) {
              for (int j = 0; j < 10000; j++) {
                Row row = sheet.getRow(i);

                if (row != null) {
                    // Lấy giá trị từ cột A (cell index 0) và B (cell index 1)
                    Cell cellA = row.getCell(0);
                    Cell cellB = row.getCell(1);

                    // Kiểm tra kiểu dữ liệu trước khi đọc giá trị
                    int a = (int)getCellValue(cellA);
                    int b = (int)getCellValue(cellB);

                    // Tính tổng
                    int sum = a + b;

                    // Tạo cell mới hoặc cập nhật giá trị cell hiện tại
                    Cell cellSum = row.createCell(2); // Cột thứ 3 (cell index 2)
                    cellSum.setCellValue(sum);
                }
              }
            }

            // Ghi dữ liệu vào file Excel
            try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
                workbook.write(outputStream);
            }

            System.out.println("Tổng của mỗi cặp số từ dòng 2 đến 10001 đã được cập nhật thành công.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static double getCellValue(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case STRING:
                    try {
                        return Double.parseDouble(cell.getStringCellValue());
                    } catch (NumberFormatException e) {
                        // Xử lý nếu giá trị không thể chuyển đổi thành số
                        return 0.0;
                    }
                default:
                    return 0.0;
            }
        }
        return 0.0;
    }
}