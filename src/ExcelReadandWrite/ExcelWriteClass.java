package ExcelReadandWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteClass {
	private static final String FileX = System.getProperty("user.dir") + "//Xcel//MyProject.xlsx";
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		  System.out.println("Writing file");
		  //ExcelWrite();
		  EnterData();
		 
		
	}
	//Write Excel
	
		public static void ExcelWrite() {
			XSSFWorkbook MyProject = new XSSFWorkbook();
			Sheet Sheet = MyProject.createSheet("My Project 2");

			Object[][] MyTable = { { "Test", "Locators", "Action","Value" }, 
					{ "open A browser", "NA", "OpenBrowser","Chrome" },
					{ "Java", "04-12-2010", "11-14-2011","NA" }, 
					{ "SQL", "09-22-2012", "12-22-2012","NA" },
					{ "SQL", "09-22-2012", "12-22-2012","NA" },
					{ "SQL", "09-22-2012", "12-22-2012","NA" },
					{ "SQL", "09-22-2012", "12-22-2012","NA" },
					{ "SQL", "09-22-2012", "12-22-2012","NA" }

			};
			int rowNum = Sheet.getLastRowNum();

			System.out.println("Creating Excel");
			for (Object[] Table : MyTable) {

				Row MyRow = Sheet.createRow(rowNum++);
				int colNum = 0;

				for (Object field : Table) {
					Cell MyCell = MyRow.createCell(colNum++);

					if (field instanceof String) {
						MyCell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						MyCell.setCellValue((Integer) field);
					}
				}
			}

			try {
				FileOutputStream EnterMyInfo = new FileOutputStream(FileX);
				MyProject.write(EnterMyInfo);
				MyProject.close();
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			System.out.println("File Closed");
		}

		public static void EnterData() {
			try {
				FileInputStream EnterMyInfo = new FileInputStream(FileX);
				XSSFWorkbook Wb= new XSSFWorkbook(EnterMyInfo);
				XSSFSheet Sh= Wb.getSheetAt(0);
				XSSFRow Row1= Sh.createRow(6);
				XSSFCell Cell1=Row1.createCell(0);
				Cell1.setCellValue("EXcel Works");
				XSSFCell Cell2=Row1.createCell(1);
				Cell2.setCellValue("01-11-2019");
				XSSFCell Cell3=Row1.createCell(2);
				Cell3.setCellValue("01-12-2019");
				XSSFCell Cell4=Row1.createCell(3);
				Cell4.setCellValue("NA");
				EnterMyInfo.close();
				FileOutputStream EndMyInfo = new FileOutputStream(FileX);
				Wb.write(EndMyInfo);
				Wb.close();
				System.out.println("Update Completed");
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch blockXf
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			
		}
}
