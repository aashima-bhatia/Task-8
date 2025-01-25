package com.files.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewTask {

		public static void main(String[] args) {
			String filepath1 = "C:\\Users\\hp\\Documents\\TESTING\\excel files\\taskdata.xlsx";
			readXLSXFile(filepath1);
			writeXLSXFile(filepath1);

		}

		private static void writeXLSXFile(String file) {
			try {
				XSSFWorkbook workbook1 = new XSSFWorkbook(new FileInputStream(file));
				XSSFSheet sheet = workbook1.getSheet("Sheet1");
				if(sheet ==null) {
					sheet = workbook1.createSheet("Sheet1");
				}
				// add new rows and columns
				int lastRowNum = sheet.getLastRowNum();
				XSSFRow newRow = sheet.createRow(lastRowNum + 1);
				
				newRow.createCell(0).setCellValue("Smith");
				newRow.createCell(1).setCellValue(32);
				newRow.createCell(2).setCellValue("smith@test1.com");
				
				// write updated workbook tp same file
				try (FileOutputStream fos = new FileOutputStream(file)){
					workbook1.write(fos);
					System.out.println("Data is written successfully.");
				}
				
			}catch(IOException e){
				System.out.println(e);
			}
			       
			
		}

		private static void readXLSXFile(String file) {
			try {
			XSSFWorkbook worknew = new XSSFWorkbook(new FileInputStream(file));
//			XSSFWorkbook write = new XSSFWorkbook(new XSSFFactory(file));
			XSSFSheet sheet = worknew.getSheet("Sheet1");
			//XSSFRow row = sheet.getRow(0);
			//XSSFRow row1 = sheet.getRow(1);
			XSSFRow row;
			int i=0;
			while((row = sheet.getRow(i))!=null) {
	                System.out.println("Name :: " + row.getCell(0).getStringCellValue());
	                System.out.println("Age :: " + row.getCell(1).getNumericCellValue());
	                System.out.println("Email Id :: " + row.getCell(2).getStringCellValue());

				i++;
			}  
			
			}catch(IOException e) {
				System.out.println(e);
			}


	}

}
