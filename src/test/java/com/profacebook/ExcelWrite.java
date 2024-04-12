package com.profacebook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public static void main(String[] args) throws IOException {
		File excelLoc = new File("C:\\Users\\gmoni\\OneDrive\\Pictures\\Face\\Excel\\New.xlsx");
		   Workbook w = new XSSFWorkbook();
		   Sheet s = w.createSheet("course" );
	         Row r = s.createRow(4);
	          Cell c = r.createCell(4);
	          c.setCellValue("selenium");
	      
	          FileOutputStream fOut = new FileOutputStream(excelLoc);
	          w.write(fOut);
	          System.out.println("finish");
	      
	          
	
	
	
	}
   
}
