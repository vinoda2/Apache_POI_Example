package HairDryer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class HairDryerInfo {

	public static void main(String[] args) throws IOException {
		File file=new File("E:\\\\Excel_Data\\\\Hairdryer_Data.xlsx");
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet=workbook.createSheet("ItemList");
		
		HSSFRow rowhead = sheet.createRow((short)0); 
		rowhead.createCell(0).setCellValue("Id");  
		rowhead.createCell(1).setCellValue("Brand");  
		rowhead.createCell(2).setCellValue("Price");  
		rowhead.createCell(3).setCellValue("Color");  
		rowhead.createCell(4).setCellValue("Availability");
		
		HSSFRow row = sheet.createRow((short)1); 
		row.createCell(0).setCellValue("X01");  
		row.createCell(1).setCellValue("Sony");  
		row.createCell(2).setCellValue("1000");  
		row.createCell(3).setCellValue("Blue");  
		row.createCell(4).setCellValue("Yes");
		
		HSSFRow row1 = sheet.createRow((short)2); 
		row1.createCell(0).setCellValue("X02");  
		row1.createCell(1).setCellValue("Mobilla");  
		row1.createCell(2).setCellValue("1500");  
		row1.createCell(3).setCellValue("White");  
		row1.createCell(4).setCellValue("Yes");
		
		HSSFRow row2 = sheet.createRow((short)3); 
		row2.createCell(0).setCellValue("X03");  
		row2.createCell(1).setCellValue("Coke");  
		row2.createCell(2).setCellValue("500");  
		row2.createCell(3).setCellValue("Black");  
		row2.createCell(4).setCellValue("yes");
		
		HSSFRow row3 = sheet.createRow((short)4); 
		row3.createCell(0).setCellValue("X04");  
		row3.createCell(1).setCellValue("Sky");  
		row3.createCell(2).setCellValue("800");  
		row3.createCell(3).setCellValue("Brown");  
		row.createCell(4).setCellValue("yes");
		
		FileOutputStream data = new FileOutputStream(file);  
		workbook.write(data);  
		
		data.close();
		workbook.close();
		
		System.out.println("Excelsheet created successfully");
		
		
	}
}
