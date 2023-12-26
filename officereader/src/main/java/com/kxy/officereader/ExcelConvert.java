package com.kxy.officereader;

import android.util.Log;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
public class ExcelConvert {

    public static String toXlsFile(String filePath) {
        String outputPath = filePath + "x";

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new HSSFWorkbook(fis)) {

            // Create a new XSSFWorkbook (xlsx) to write the data
            try (Workbook newWorkbook = new XSSFWorkbook()) {
                Sheet oldSheet = workbook.getSheetAt(0);
                Sheet newSheet = newWorkbook.createSheet("Sheet1");

                // Copy data from old sheet to new sheet
                copySheet(oldSheet, newSheet);

                // Save the new workbook to a file
                try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                    newWorkbook.write(fos);
                }
            }

            System.out.println("Conversion completed successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }

//        try {
//            // 1. 读取旧的xls文件
//            FileInputStream xlsFile = new FileInputStream(filePath);
//            Workbook workbook = new HSSFWorkbook(xlsFile);
//
//            // 2. 创建新的xlsx文件
//            Workbook newWorkbook = new XSSFWorkbook();
//
//            // 3. 复制数据到新的xlsx文件
//            copyData(workbook, newWorkbook);
//
//            // 4. 保存新的xlsx文件
//            FileOutputStream xlsxFile = new FileOutputStream(outputPath);
//            newWorkbook.write(xlsxFile);
//
//            // 5. 关闭文件流
//            xlsFile.close();
//            xlsxFile.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
        return outputPath;
    }

    private static void copyData(Workbook sourceWorkbook, Workbook targetWorkbook) {
        for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {
            Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
            Sheet targetSheet = targetWorkbook.createSheet(sourceSheet.getSheetName());

            for (int j = 0; j < sourceSheet.getPhysicalNumberOfRows(); j++) {
                Row sourceRow = sourceSheet.getRow(j);
                Row targetRow = targetSheet.createRow(j);

                for (int k = 0; k < sourceRow.getPhysicalNumberOfCells(); k++) {
                    Cell sourceCell = sourceRow.getCell(k);
                    Cell targetCell = targetRow.createCell(k);

                    // 复制单元格的内容和样式
                    copyCell(sourceCell, targetCell, targetWorkbook);
                }
            }
        }
    }

    private static void copyCell(Cell sourceCell, Cell targetCell, Workbook targetWorkbook) {
        if (sourceCell != null) {
            CellStyle sourceCellStyle = sourceCell.getCellStyle();
            Log.e("测试样式", sourceCell.getCellStyle().getClass().toString());
            CellStyle targetCellStyle =  targetWorkbook.createCellStyle();
            Log.e("测试样式2", targetWorkbook.createCellStyle().getClass().toString());

            // 复制常规样式
            targetCellStyle.cloneStyleFrom(sourceCellStyle);

            // 复制特定样式（如果有）
            if (sourceCellStyle instanceof XSSFCellStyle && targetCellStyle instanceof XSSFCellStyle) {
                XSSFCellStyle sourceXSSFCellStyle = (XSSFCellStyle) sourceCellStyle;
                XSSFCellStyle targetXSSFCellStyle = (XSSFCellStyle) targetCellStyle;

                // 复制XSSFCellStyle特有的属性
                targetXSSFCellStyle.setFillBackgroundColor(sourceXSSFCellStyle.getFillBackgroundColor());
                // ... 复制其他属性
            }

            targetCell.setCellStyle(targetCellStyle);


            switch (sourceCell.getCellType()) {
                case BLANK:
                    break;
                case BOOLEAN:
                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case ERROR:
                    targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                    break;
                case FORMULA:
                    targetCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                case NUMERIC:
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case STRING:
                    targetCell.setCellValue(sourceCell.getStringCellValue());
                    break;
            }
        }
    }


    private static void copySheet(Sheet oldSheet, Sheet newSheet) {
        for (int i = 0; i < oldSheet.getPhysicalNumberOfRows(); i++) {
            Row oldRow = oldSheet.getRow(i);
            Row newRow = newSheet.createRow(i);

            if (oldRow != null) {
                for (int j = 0; j < oldRow.getPhysicalNumberOfCells(); j++) {
                    Cell oldCell = oldRow.getCell(j);
                    Cell newCell = newRow.createCell(j);

                    if (oldCell != null) {
                        // Copy cell style and value
                        newCell.setCellStyle(oldCell.getCellStyle());
                        switch (oldCell.getCellType()) {
                            case STRING:
                                newCell.setCellValue(oldCell.getStringCellValue());
                                break;
                            case NUMERIC:
                                newCell.setCellValue(oldCell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                newCell.setCellValue(oldCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                newCell.setCellFormula(oldCell.getCellFormula());
                                break;
                            default:
                                // Do nothing
                        }
                    }
                }
            }
        }
    }
}
