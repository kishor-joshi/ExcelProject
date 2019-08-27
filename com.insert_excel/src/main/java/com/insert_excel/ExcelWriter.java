package com.insert_excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.read.biff.BiffException;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

public class ExcelWriter {

	private static String[] columns = { "Name", "Email", "UserId", "Salary" };
	private static List<Employee> employees = new ArrayList<>();

	// Initializing employees data to insert into the excel file
	public static void Create(Workbook workbook, Sheet sheet) throws InvalidFormatException, IOException {

		Scanner input = new Scanner(System.in);
		System.out.println("Enter Employee Details");
		System.out.println("Enter Employee Name");
		String name = input.next();
		System.out.println("Enter Employee Email-Id");
		String emailId = input.next();
		System.out.println("Enter Employee Id");
		int id = input.nextInt();
		System.out.println("Enter Employee Salary");
		double salary = input.nextDouble();

		employees.add(new Employee(name, emailId, id, salary));

		// Create a Font for styling header cells
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		// Create a CellStyle with the font
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Create cells
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);
		}

		// Create Cell Style for formatting Date
		CellStyle dateCellStyle = workbook.createCellStyle();

		// Create Other rows and cells with employees data
		int rowNum = 1;
		for (Employee employee : employees) {
			Row row = sheet.createRow(rowNum++);

			row.createCell(0).setCellValue(employee.getName());

			row.createCell(1).setCellValue(employee.getEmail());

			row.createCell(2).setCellValue(employee.getuserId());

			row.createCell(3).setCellValue(employee.getSalary());
		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);
		}

		// Write the output to a file
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream("D:\\java_training\\com.insert_excel\\poi-generated-file.xlsx");
			workbook.write(fileOut);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Menu m=new Menu(workbook,sheet);
		fileOut.close();
	}

	public static void main(String[] args) throws InvalidFormatException, IOException  {
		
		File file=new File("D:\\java_training\\com.insert_excel\\poi-generated-file.xlsx");
		
		FileOutputStream fileOut = new FileOutputStream(file);
		Workbook workbook = new XSSFWorkbook();
		// Create a Sheet
		Sheet sheet = workbook.createSheet("Employee");

		
        if(!file.exists())
        {
        	
    		System.out.println("file created with workbook");
    		Menu m=new Menu(workbook,sheet);
        }
        else
        {
        	
        	Menu m=new Menu(workbook,sheet);
        	
        }
        
		

	}

}
