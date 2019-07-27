package com.honwaii.excel.exceloperation.operation;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class ExcelFilesOperation {

    public static void main(String[] args) throws IOException {
        List<String> files = getAllExcelFiles();
        for (String file : files) {
            System.out.println("更新文件:" + file);
            readAndWriteContent(file);
        }
    }

    private static void readAndWriteContent(String path) throws IOException {
        File xlsFile = new File(path);
        FileInputStream fis = new FileInputStream(xlsFile);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheet("Sheet1");
        int rows = sheet.getLastRowNum();
        Row r = sheet.getRow(3);
        int spec = 0;
        int st = 0;
        int updateCol = 0;
        for (int i = 0; i < r.getLastCellNum(); i++) {
            if ("Specification".equalsIgnoreCase(r.getCell(i).getStringCellValue())) {
                spec = i;
            }
            if ("Standard".equalsIgnoreCase(r.getCell(i).getStringCellValue())) {
                st = i;
            }
            if (r.getCell(i).getStringCellValue().contains("图号")) {
                updateCol = i;
            }
        }

        for (int i = 5; i <= rows; i++) {
            Cell t = sheet.getRow(i).getCell(spec);
            if (t == null) {
                continue;
            }
            String specifiction = t.getStringCellValue();
            String standard = sheet.getRow(i).getCell(st).getStringCellValue();
            String temp = standard + " " + specifiction.replace(" ", "");
            sheet.getRow(i).getCell(updateCol).setCellValue(temp);
            if (i == 5) {
                System.out.println(temp + " -> rows=" + rows);
            }
        }
        fis.close();
        FileOutputStream xlsStream = new FileOutputStream(xlsFile);
        workbook.write(xlsStream);
        workbook.close();
    }


    private static List<String> getAllExcelFiles() {
        String path = "E:\\solidworks\\toolbox\\datas";
        File file = new File(path);
        return func(file);
    }

    private static List<String> func(File file) {
        File[] files = file.listFiles();
        List<String> excelFiles = new ArrayList<>();
        if (files == null) {
            return excelFiles;
        }
        for (File f : files) {
            if (f.isDirectory()) {
                List<String> temp = func(f);
                if (temp != null) {
                    excelFiles.addAll(func(f));
                }
                continue;
            }
            excelFiles.add(f.getAbsolutePath());
        }
        return excelFiles;
    }
}
