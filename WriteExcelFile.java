package com.ExcelSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile {
	
	
	//create a method to write a Excel File(Create)

	public static void writeCountryListToFile(String fileName,List<Country> countryList) throws Exception ,IOException{
		
		//create a workbook
		
		Workbook workbook =null;
		
		if(fileName.endsWith("xlsx")) {
			
			  workbook =  new  XSSFWorkbook();
			
		}else if(fileName.endsWith("xls")) {
			
			 workbook = new HSSFWorkbook();
		}else {
			
			throw new Exception("file not found !, please check file name xls or xlsx ");
		}
		
	         Sheet sheet =	workbook.createSheet("countries");
		
		    Iterator<Country> itr = countryList.iterator();
		    
		    int rowIndex =0;
		    while(itr.hasNext()) {
		    	
		    	  Country country = itr.next();
		    	  
		    	  Row row = sheet.createRow(rowIndex++);
                  Cell cell0 = row.createCell(0);
                       cell0.setCellValue(country.getName());
                  Cell cell1 = row.createCell(1);
                       cell1.setCellValue(country.getShortCode());
                             
		    }
		    
		    //let's now write a excel data to file now.
		    
		    FileOutputStream fos = new FileOutputStream(fileName);
		    workbook.write(fos);
		    
		    fos.close();
		    System.out.println(fileName +" Successfully Write...");
	}
	
	//main method
	public static void main(String[] args) throws Exception{
		
		List<Country> list = ReadExcelFileToList.readExcelData("C:/Educative.io/Aa_Lets Start/1111/1.IMP/Java for Programmers/5. Java Collections/11. Arrays/Temporary.xlsx");
		
		
		WriteExcelFile.writeCountryListToFile("C:/Educative.io/Aa_Lets Start/1111/1.IMP/Java for Programmers/5. Java Collections/11. Arrays/CountryList.xls", list);
	}

}
