package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Workbook --> Sheet --> Rows --> Cells

public class WritingToExcel2 {

	
	public static void main(String[] args) throws IOException {
		
		//To write data, we can Collections --> object Array, ArrayList, HashMap
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Emp Info");
		
		ArrayList<Object[]> empdata=new ArrayList<Object[]>(); // every object is single dimensional array
		
		empdata.add(new Object[]{"EmpID", "Name", "Job"});
		empdata.add(new Object[]{101,"David","Engineer"});
		empdata.add(new Object[]{102,"Smith","Manager"});
		empdata.add(new Object[]{103,"Scott","Analyst"});
		
		// Using for..each loop
		
		int rowCount=0;
		
		for(Object[] emp:empdata)
		{
			XSSFRow row=sheet.createRow(rowCount++);
			int columnCount=0;
			for(Object value:emp)
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
