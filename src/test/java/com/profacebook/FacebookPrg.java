package com.profacebook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class FacebookPrg {

	public static void main(String[] args) throws IOException {
		File excelLoc = new File("C:\\Users\\gmoni\\OneDrive\\Pictures\\Face\\Excel\\moninewexcel.xlsx");
		
		FileInputStream f = new FileInputStream(excelLoc);
		
		Workbook w = new XSSFWorkbook(f);
		
		Sheet s = w.getSheet("Sheet1");
		
		Row r = s.getRow(1);
		
		Cell c = r.getCell(1);
		
		System.out.println(c);
		
		int rows = s.getPhysicalNumberOfRows();
		System.out.println(rows);
		
		int cell = r.getPhysicalNumberOfCells();
		System.out.println(cell);
		
		
		
		
		
				}

}
