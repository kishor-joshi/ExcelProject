package com.insert_excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelDelete {

	public void delete(Workbook workbook, Sheet sheet) {
		// TODO Auto-generated method stub
		
		 //Find number of rows in excel file
        Iterator<Row> rowIterator = sheet.iterator();
        Scanner input=new Scanner(System.in);
        System.out.println("enter the name whose data is to be deleted");
        String name=input.next();
        int rowNum=0;
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    if (cell.getRichStringCellValue().getString().trim().equals(name)) 
                    {
                         rowNum=row.getRowNum();
                                                 
                    }
                }
            }
        } 
        System.out.println(rowNum+" is deleted");
        sheet.removeRow(sheet.getRow(rowNum));
        FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream("D:\\java_training\\com.insert_excel\\poi-generated-file.xlsx");
			workbook.write(fileOut);
			Menu m=new Menu(workbook,sheet);
			fileOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
      
		
	}

}
