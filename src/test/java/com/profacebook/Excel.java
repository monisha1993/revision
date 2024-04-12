package com.profacebook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
public static void main(String[] args) throws IOException {
	File excelLoc = new File("C:\\Users\\gmoni\\OneDrive\\Pictures\\Face\\Excel\\Adactin.xlsx"); 
	FileInputStream fIn = new FileInputStream(excelLoc);

	Workbook w = new XSSFWorkbook(fIn);
	
	Sheet s = w.getSheet("sheet1");
	
	
	
	 for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
	   Row r = s.getRow(i);
		 for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
			Cell c = r.getCell(j);
		System.out.println(c);
		 int type = c.getCellType();
		 if (type==1) {
			String cellValue = c.getStringCellValue();
			System.out.println(cellValue);
		}
		 else if (type==0) {
			if (DateUtil.isCellDateFormatted(c)) {
				Date d= c.getDateCellValue();
				SimpleDateFormat sdf = new SimpleDateFormat("dd/mm/yy");
				String f = sdf.format(d);
			  System.out.println(f);
			}
			else {
				double n = c.getNumericCellValue();
			long l= (long)n;
			String valueOf = String.valueOf(l);		
			System.out.println(valueOf);
			
			
			
			
			}
			
		}
			
		 
		 
		 }
		 
		 
	}
	
	
	
	
	
	
	
	
	
	
		}
	
	
	
	
	
	
	

}

