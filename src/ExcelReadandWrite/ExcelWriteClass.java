package ExcelReadandWrite;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteClass {
	private static final String FileX = System.getProperty("user.dir") + "//Xcel//MyProject.xlsx";
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		  System.out.println("Writing file");
		  ExcelWrite();
		 
		
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

}
