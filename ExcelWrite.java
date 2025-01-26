package excelReadWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {
	public static void main(String[] args) {
		ExcelWrite x = new ExcelWrite();

		for (int i = 0; i < 5; i++) {
			for (int j = 0; j <3; j++) {
				
				System.out.print(x.getExcelData("Sheet2", i, j)+ " ");
			}
			System.out.println("  ");
		}
		x.WriteToExcel("Sheet2",0, 3, "Result");
		x.WriteToExcel("Sheet2",1, 3, "Fail");
		x.WriteToExcel("Sheet2",2, 3, "Pass");
		x.WriteToExcel("Sheet2",3, 3, "Fail");
		x.WriteToExcel("Sheet2",4, 3, "Fail");
	

	}
	public void WriteToExcel(String Sheetname, int rowNum, int cellNum, String desc) {
		FileInputStream fis;
		XSSFWorkbook wb;
		
		try {
			fis = new FileInputStream("Utils//Students.xlsx");
			wb=new XSSFWorkbook(fis);
			XSSFSheet s = wb.getSheet(Sheetname);
			XSSFRow r = s.getRow(rowNum);
			XSSFCell c=r.createCell(cellNum);
			c.setCellValue(desc);
			FileOutputStream fos=new FileOutputStream("Utils//Students.xlsx");
			wb.write(fos);
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		}catch(IOException e) {
			
			e.printStackTrace();
		}
			
		}
		
			
	public String getExcelData(String sheetName, int rowNum, int ColNum) {
		String retVal = null;
		try {
			FileInputStream fis = new FileInputStream("Utils//Students.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet s = wb.getSheet(sheetName);
			XSSFRow r = s.getRow(rowNum);
			XSSFCell c = r.getCell(ColNum);
			retVal = ExcelWrite.getCellValue(c);
			fis.close();
			wb.close();

		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return retVal;
	}
	
	public static String getCellValue(XSSFCell c) {
		switch(c.getCellType()) {
		case NUMERIC:
			return String.valueOf(c.getNumericCellValue());
		case BOOLEAN:
			return String.valueOf(c.getBooleanCellValue());
		case STRING:
			return c.getStringCellValue();
			default:
				return c.getStringCellValue();	
		}
	}
	}
/*Output:
Name Age Email   
Ram 10.0 r@x.com   
Madhan 20.0 Mad@m.com   
Lalith 25.0 lal@l.com   
Vimal 36.0 vim@v.com   */