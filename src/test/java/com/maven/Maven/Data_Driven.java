package com.maven.Maven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Data_Driven {

	public static void main(String[] args)  throws IOException{
		/*File f = new File("C:\\Users\\721901\\eclipse-workspace\\Maven\\target\\Credentials.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell = row.getCell(j);
				CellType cellType = cell.getCellType();
				if (cellType.equals(cellType.STRING)) {
					System.out.println(cell.getStringCellValue());
				}else if (cellType.equals(cellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					System.out.println(String.valueOf(numericCellValue));
				}
			}
				
			}*/
		
		File f = new File("C:\\Users\\721901\\eclipse-workspace\\Maven\\target\\Credentials.xlsx");
		FileInputStream fis = new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		Sheet sheet = (Sheet) wb.createSheet("Register");
		Row createRow=((org.apache.poi.ss.usermodel.Sheet) sheet).createRow(0);
		Cell createCell = createRow.createCell(0);
		createCell.setCellValue("25656");
		
		FileOutputStream fos = new FileOutputStream(f);
		wb.write(fos);
		wb.close();
		System.out.println("Completed");
		}

	}


