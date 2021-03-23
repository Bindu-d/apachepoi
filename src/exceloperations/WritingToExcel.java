package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Workbook --> Sheet --> Rows --> Cells

public class WritingToExcel {

	
	public static void main(String[] args) throws IOException {
		
		//To write data, we can Collections --> object Array, ArrayList, HashMap
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Emp Info");
		
		Object empdata[][]= {	{"EmpID", "Name", "Job"},
								{101,"David","Engineer"},
								{102,"Smith","Manager"},
								{103,"Scott","Analyst"}
							};
//		// Using for loop
//		int rows= empdata.length;
//		int cols=empdata[0].length;
//		
//		System.out.println(rows); //4
//		System.out.println(cols); //3
//		
//		for(int r=0; r<rows;r++)
//		{
//			XSSFRow row= sheet.createRow(r); //0
//			
//			for(int c=0;c<cols;c++)
//			{
//				XSSFCell cell=row.createCell(c); //0
//				Object value=empdata[r][c];  //0 0
//				
//				if(value instanceof String)
//					cell.setCellValue((String)value);
//				if(value instanceof Integer)
//					cell.setCellValue((Integer)value);
//				if(value instanceof Boolean)
//					cell.setCellValue((Boolean)value);				
//			}
//		}
		
		
		// Using for..each loop
		
		int rowCount=0;
		
		for (Object emp[]:empdata) //It gets the entire row or 1st record
		{
			XSSFRow row= sheet.createRow(rowCount++);
			int columnCount=0;
			for(Object value:emp) //we have 3 values in emp
			{
				XSSFCell cell=row.createCell(columnCount++);
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			}
		}
		
		String filePath=".\\datafiles\\employee.xlsx";
		
		FileOutputStream outstream=new FileOutputStream(filePath);
		
		workbook.write(outstream);
		
		outstream.close();
		
		System.out.println("employee.xlsx file has been written successfully..");
		
	}

}
