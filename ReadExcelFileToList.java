package com.ExcelSheet;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFileToList {

	
	//creare a method to read excel data
	
	public static List<Country> readExcelData(String fileName) throws FileNotFoundException,IOException{
		
		
		List<Country> countriesList =  new ArrayList<Country>();
		
		//Create the input stream from the xlsx/xls file
		try {
			FileInputStream fis = new FileInputStream(fileName);
			
			//Create Workbook instance for xlsx/xls file input stream
			Workbook workbook = null;
			
			if(fileName.toLowerCase().endsWith("xlsx")) {
				
				workbook = new XSSFWorkbook(fis);
			}else if(fileName.toLowerCase().endsWith("xls")) {
				
				workbook = new XSSFWorkbook(fis);
			}
			
			//Get the number of sheets in the xlsx file
			 int noOfSheets = workbook.getNumberOfSheets();
			
			
			//loop through each of the sheets
			for(int i=0; i<noOfSheets;i++) {
				
				//Get the nth sheet from the workbook
				
				Sheet sheet =  workbook.getSheetAt(i);
				
				//every sheets has a row, iterate over them
				
			  Iterator<Row> 	rowIterator = sheet.iterator();
			  
			  while(rowIterator.hasNext()) {
				  
				     String name ="";
				     String shortCode="";
			
				     Row row = rowIterator.next();
				     
				     //every row has a columns, so iterate over them
				    
				     Iterator<Cell> cellIterator = row.iterator();
				     
				     while(cellIterator.hasNext()) {
				    	 
				    	 //get the cell Object
				    	Cell cell = cellIterator.next();
				    	
				    	
				    	//check the cell type and process accordingly
				    	
				    	switch(cell.getCellType()){
				    		
				    	case Cell.CELL_TYPE_STRING:
				    		 
				    		if(shortCode.equalsIgnoreCase("")) {
				    			
				    			  shortCode = cell.getStringCellValue().trim();
				    		}
				    		else if(name.equalsIgnoreCase("")) {
				    			
				    			   name = cell.getStringCellValue();
				    		}else {
				    			
				    			//random data leave it
				    			System.out.println("Random data :: " + cell.getStringCellValue());
				    		}
				    		
				    		break;
				    		
				    	case Cell.CELL_TYPE_NUMERIC:
				    	
				    		System.out.println("Random data:: " +cell.getNumericCellValue());
				    	
				    	}
				     } //end of cell Iterator
				     
				       Country c= new Country(name, shortCode);
				         countriesList.add(c);
				     
			  }//end of row Iterator
			} //end of sheet for loop
			
			 //close file input stream
			fis.close();
			 
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		}
		
		return countriesList;
		
		
		
	}
	
	
	//create a main Method
	
	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		List<Country> list = readExcelData("C:/Educative.io/Aa_Lets Start/1111/1.IMP/Java for Programmers/5. Java Collections/11. Arrays/Temporary.xlsx");
		System.out.println("Country List\n"+list);
	}
}
