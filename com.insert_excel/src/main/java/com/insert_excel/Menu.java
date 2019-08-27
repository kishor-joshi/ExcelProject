package com.insert_excel;

import java.io.IOException;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Menu {
	Menu(Workbook workbook, Sheet sheet) throws InvalidFormatException, IOException{
	
	Scanner input = new Scanner(System.in);
	System.out.println();
	System.out.println("enter 1:Create  Employee Details");
	System.out.println("enter 2:Fetch  Employee Details");
	System.out.println("enter 3:Delete  Employee Details");
	System.out.println("enter 4:Update  Employee Details");	
	System.out.println("enter any other to exit");
	int choice = input.nextInt();

	switch (choice) {
	case 1:
		ExcelWriter excelwriter = new ExcelWriter();
		excelwriter.Create(workbook,sheet);
		break;
	case 2:
		ExcelReader excelreader = new ExcelReader();

		excelreader.SearchExistingWorkbook(workbook,sheet);
		break;
	
	 case 3:
		 ExcelDelete exceldelete =new ExcelDelete(); 
		 exceldelete.delete(workbook,sheet);
	     break; 
	 case 4:
		 ExcelUpdate excelupdate =new ExcelUpdate(); 
		 excelupdate.update(workbook,sheet);
	     break;
	
	 
	default:
		System.exit(0);

	}
	}
}
