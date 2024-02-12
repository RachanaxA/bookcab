package utilities;

import java.io.FileOutputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelutilities2

{

	public void writeAvailableBookings(int len, String name) throws IOException

	{

		FileOutputStream file1 = new FileOutputStream(
				"C:\\Users\\2303455\\eclipse-workspace\\Hackathon\\Excel\\excel.xlsx");

		XSSFWorkbook wb = new XSSFWorkbook();

		XSSFSheet sheet1 = wb.createSheet("Lowest Price");
	

		int totalRow = len;

		for (int i = 0; i <= totalRow; i++) {

			Row row = sheet1.createRow(i);

			Cell co1 = row.createCell(0);

			co1.setCellValue(name);

		}

		wb.write(file1);

		wb.close();

		file1.close();

		System.out.println("Excel Writing is Done(Profile Info)!!!!!");

	}

}