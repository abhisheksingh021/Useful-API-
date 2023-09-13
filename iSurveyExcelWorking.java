package com.ExcelSheet;

import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class iSurveyExcelWorking {
	
	public void downloadResponsesData(String fileName)
			throws IOException {
//		System.out.println("Request Start "+new Date());
		/*
		 * String surveyid = request.getParameter("surveyid"); String surveyName =
		 * request.getParameter("surveyName"); String langid =
		 * request.getParameter("languageid"); String
		 * downLoadDate=request.getParameter("downLoadDate");
		 * downLoadDate=downLoadDate==null?"1900-01-01":downLoadDate;
		 */
		
		List<List<String>>  survlist= surveyMstLogic.downloadResponsesData(surveyid, langid,downLoadDate);
//		System.out.println("Excel generation Started "+new Date());
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("User Responses");
		CellStyle opencellstyleHead = workbook.createCellStyle();
		HSSFFont font = (HSSFFont) workbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		opencellstyleHead.setFont(font);  
		opencellstyleHead.setWrapText(true);
		
		CellStyle style = workbook.createCellStyle();
		style.setWrapText(true);
		
		CellStyle styleYellow = workbook.createCellStyle();
		
		styleYellow.setWrapText(true);
		styleYellow.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		//styleYellow.setFillBackgroundColor(HSSFColor.GOLD.index);
		styleYellow.setBorderBottom(HSSFCellStyle.BORDER_HAIR);
		styleYellow.setBorderTop(HSSFCellStyle.BORDER_HAIR);
		styleYellow.setBorderRight(HSSFCellStyle.BORDER_HAIR);
		styleYellow.setBorderLeft(HSSFCellStyle.BORDER_HAIR); 
		//styleYellow.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
		styleYellow.setFillPattern(CellStyle.SOLID_FOREGROUND); 
		Font fontYellow  = workbook.createFont();
        fontYellow.setColor(IndexedColors.BLACK.getIndex());
        styleYellow.setFont(fontYellow);
		
		
		Row row=null;
		Cell cell =null;
		
		int i=0;
		int colsize=0;
		int fn=1;
		while(i<survlist.size()){
			row = sheet.createRow((short) i);
			List<String> innerList=survlist.get(i);
			if(i==0){
				colsize=innerList.size();
				int j=1;
				int x=1;
				while(j<innerList.size()){
					cell = row.createCell(j-x);
					cell.setCellValue(innerList.get(j));
					cell.setCellStyle(opencellstyleHead);
					j++;
					if(j>fn && j<innerList.size())
					{j++;x++;}
				}
			}else{
				int j=1;
				int x=1;
				while(j<innerList.size()){
					cell = row.createCell(j-x);
					cell.setCellValue(innerList.get(j));
					cell.setCellStyle(style);
					j++;
					if(j>fn && j<innerList.size())
						{
							String val=innerList.get(j);
							if(val.trim().equalsIgnoreCase("Y"))
							{
								cell.setCellStyle(styleYellow);
							}
							j++;
							x++;
						}
				}
			}
			
			i++;
		}
		for (i = 0; i < colsize; i++) {
			  if(i<=4){
				   sheet.autoSizeColumn(i);   
			   }else{
				   sheet.setColumnWidth(i, 6000);	
			   }
		}
//		System.out.println("Excel generated "+new Date());
		response.setCharacterEncoding("UTF-8");
		response.setContentType("application/vnd.ms-excel");
		String fileName = surveyName+"-Responses.xls";
		
		response.setHeader("Content-Disposition", "attachment ;filename="
				+ fileName);
		response.addHeader("Cache-Control", "max-age=1, must-revalidate");
		
		try {
			workbook.write(response.getOutputStream());
//			System.out.println("Resoponse generated "+new Date());
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}

}
