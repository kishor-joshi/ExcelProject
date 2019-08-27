package com.insert_excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUpdate {
	public static void update(Workbook workbook,Sheet sheet) throws IOException {
	Scanner s=new Scanner(System.in);
	System.out.println("enter emp name whose value has to change");
	String old=s.next();
	String converttoString="";

	        try {
	            FileInputStream file = new FileInputStream("D:\\java_training\\com.insert_excel\\poi-generated-file.xlsx");
	            int row=sheet.getLastRowNum();
	            Cell cell = null;
	          //Retrieve the row and check for null
	            boolean b=true;
	            for(int r=1;r<=row;r++) {
	            
	            XSSFRow sheetrow = (XSSFRow) sheet.getRow(r);
	            if(sheetrow == null){
	                sheetrow = (XSSFRow) sheet.createRow(r);
	            }
	            //Update the value of cell
	            cell = sheetrow.getCell(0);
	            
	            
	            converttoString=""+cell;
	            
	            
	            if(old.equals(converttoString)) {
	            //cell.setCellValue(update);
	            
	       System.out.println("Enter updated Employee email-id");
	       sheetrow.getCell(1).setCellValue(s.next());
	       System.out.println("Enter updated Employee emp id");
	       sheetrow.getCell(2).setCellValue(s.next());
	       System.out.println("Enter updated Employee SALARY");
	       sheetrow.getCell(3).setCellValue(s.next());
	    
	       b=false;
	       break;    
	            
	            }
	        }
	            if(b)
	            System.out.println("no emp name matching");
	            
	            

	            FileOutputStream outFile =new FileOutputStream("D:\\java_training\\com.insert_excel\\poi-generated-file.xlsx");
	            Menu m=new Menu(workbook,sheet);
	            workbook.write(outFile);
	            outFile.close();

	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        } catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
	    }

	
	
	
}
