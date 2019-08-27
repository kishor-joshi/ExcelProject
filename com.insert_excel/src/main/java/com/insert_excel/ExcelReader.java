package com.insert_excel;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelReader {
	 static void SearchExistingWorkbook(Workbook workbook,Sheet sheet) throws InvalidFormatException, IOException {
	       	 
		    //Find number of rows in excel file
	        Iterator<Row> rowIterator = sheet.iterator();
	        while (rowIterator.hasNext())
	        {
	            Row row = rowIterator.next();
	            //For each row, iterate through all the columns
	            Iterator<Cell> cellIterator = row.cellIterator();
	          
	             
	            while (cellIterator.hasNext())
	            {
	            	
	                Cell cell = cellIterator.next();
	               
	                //Check the cell type and format accordingly
	                switch (cell.getCellType())
	                {
	                    case Cell.CELL_TYPE_NUMERIC:
	                        System.out.print(cell.getNumericCellValue() + "\t");
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                        System.out.print(cell.getStringCellValue() + "\t");
	                        break;
	                
	                }
	            }
	            System.out.println("");
	        }
	      
	   	 Menu m=new Menu(workbook,sheet);
	    }

}
