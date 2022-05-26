package org.maven.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
public static void main(String[] args) throws IOException {
	File location=new File("D:\\Practice Files\\excel\\Excel Read Practice.xlsx");
	
	FileInputStream fin=new FileInputStream(location);
	
	Workbook w=new XSSFWorkbook(fin);
	
	Sheet sheet = w.getSheet("Sheet1");
	
	
	for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
		Row r = sheet.getRow(i);
	for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
		
		Cell c = r.getCell(j);
		int cellType = c.getCellType();
        		
		
     if (cellType==1) {
		
    	 String value = c.getStringCellValue();
    	 System.out.println(value);
    	 }
     if (cellType==0) {
		
    	double d = c.getNumericCellValue();
    	long l=(long) d;
    	String valueOf = String.valueOf(l);
    	System.out.println(valueOf);
    	
    }
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
	}	
	}
	
	
	
	
	
	
	
	
	
}
}
