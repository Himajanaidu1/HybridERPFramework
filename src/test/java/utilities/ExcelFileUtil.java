package utilities;

import java.awt.image.IndexColorModel;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelFileUtil {
Workbook wb;
//constructor for reading excel path
public ExcelFileUtil(String ExcelPath)throws Throwable
{
	FileInputStream fi=new FileInputStream(ExcelPath);
	wb=WorkbookFactory.create(fi);	
}
//count no of rows in a sheet
public int rowCount(String SheetName)
{
	return wb.getSheet(SheetName).getLastRowNum();
}
//,ethod for reading cell data
public String getCellData(String SheetName,int row,int column)
{
	String data;
	if (wb.getSheet(SheetName).getRow(row).getCell(column).getCellType()==CellType.NUMERIC)
	{
	int celldata=(int) wb.getSheet(SheetName).getRow(row).getCell(column).getNumericCellValue()	;
	data=String.valueOf(celldata);
}
else
{
data= wb.getSheet(SheetName).getRow(row).getCell(column).getStringCellValue();
}
	return data;	
}
//method foe setcelldata
public void setcellData(String sheetName,int row,int cloumns,String status,String writeExcelpath) throws Throwable
{
	//get sheet from wb
	Sheet ws=wb.getSheet(sheetName);
	//get row from sheet
	Row rowNum=ws.getRow(row);
	//Create cell in row
	Cell cell=rowNum.createCell(cloumns);
	//write status
	cell.setCellValue(status);
	if (status.equalsIgnoreCase("pass"))
	{
		CellStyle style=wb.createCellStyle();
		Font font=wb.createFont();
		//colour with green
		font.setColor(IndexedColors.GREEN.getIndex());
		font.setBold(true);
		style.setFont(font);
		ws.getRow(row).getCell(cloumns).setCellStyle(style);
	}
	else if(status.equalsIgnoreCase("fail"))
	{

		CellStyle style=wb.createCellStyle();
		Font font=wb.createFont();
		//colour with red
		font.setColor(IndexedColors.RED.getIndex());
		font.setBold(true);
		style.setFont(font);
		ws.getRow(row).getCell(cloumns).setCellStyle(style);
	}
	else if(status.equalsIgnoreCase("Blocked"))
	{

		CellStyle style=wb.createCellStyle();
		Font font=wb.createFont();
		//colour with blue
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBold(true);
		style.setFont(font);
		ws.getRow(row).getCell(cloumns).setCellStyle(style);
	}
		FileOutputStream fo=new FileOutputStream(writeExcelpath);
		wb.write(fo);
}
public static void main(String[] args)throws Throwable
{
	//create obk=ject for class
	ExcelFileUtil xl=new ExcelFileUtil("D:/Sample.xlsx");
	//count no of rows in emp sheet
	int rc=xl.rowCount("emp");
	System.out.println(rc);
	 for(int i=1;i<=rc;i++)
	 {
		 //read all rows each cell data
		 String fname=xl.getCellData("emp", rc, 0);
		 String mname=xl.getCellData("emp", rc, 1);
		 String lname=xl.getCellData("emp", rc, 2);
		 String eid=xl.getCellData("emp", rc, 3);
		 System.out.println(fname+"   "+mname+"    "+lname+"   "+eid);
		 //write as pass into status cell
	     xl.setcellData("emp", i, 4, "pass", "D:/Results.xlsx");	 
	 }
		 
}
}
