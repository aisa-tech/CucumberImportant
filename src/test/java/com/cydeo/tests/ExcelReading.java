package com.cydeo.tests;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReading {
@Test
public void  read_from_excel_file() throws IOException {
    String path = "SampleData.xlsx";
    File file = new File(path);
    //to read from excel
    FileInputStream fileInputStream = new FileInputStream(file);

    //workbook>sheet>row>cell

    //create workbook
    XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
    //we need to get specific sheet from currently opned workbook
    XSSFSheet sheet = workbook.getSheet("Employees");

    //select row and cell
    //print mary cell
    System.out.println(sheet.getRow(1).getCell(0));

    //print out developer
    System.out.println(sheet.getRow(3).getCell(2));

    //return the count of cells only
    //starts counting from 1
    int usedRows = sheet.getPhysicalNumberOfRows();
    System.out.println(usedRows);

    //returns  the number top cell to bottom cell
    //it doesnt care if the cell is empty or not
    //starts from counting from 0
    int lastUsedRow = sheet.getLastRowNum();
    System.out.println(lastUsedRow);

    //todo: Create a logic to print Vinods name
    for (int rowNum = 0; rowNum < usedRows; rowNum++) {
        if (sheet.getRow(rowNum).getCell(0).toString().equals("Vinod")){
            System.out.println(sheet.getRow(rowNum).getCell(0));
        }
    }
    //todo : create a logic to print lindas job id

    for (int rowNum = 0; rowNum < usedRows; rowNum++) {
        if (sheet.getRow(rowNum).getCell(0).toString().equals("Linda")){
            System.out.println(sheet.getRow(rowNum).getCell(2));
        }
    }






}
}
