package ExcelReadandWrite;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReadFile {

	//Describe file or path
	private static final String FileX = System.getProperty("user.dir") + "//Xcel//MyProject.xlsx";

	public static void main(String[] args) {
		
		System.out.println("Reading file");
		ReadExcelFile();
	}

	
	
	
	//read excel
	public static void ReadExcelFile() {
		try {
			FileInputStream ReadFile = new FileInputStream(FileX);
			Workbook workbook = new XSSFWorkbook(ReadFile);
			Sheet MyProjectSheet = workbook.getSheetAt(0);

			Iterator<Row> iterator = MyProjectSheet.iterator();
			while (iterator.hasNext()) {
				Row currentrow = iterator.next();
				Iterator<Cell> cellIterator = currentrow.iterator();
				while (cellIterator.hasNext()) {
					Cell currentCell = cellIterator.next();
					if (currentCell.getCellTypeEnum() == CellType.STRING) {
						System.out.println(currentCell.getStringCellValue() + "---");
					} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
						System.out.println(currentCell.getNumericCellValue() + "---");
					}
				}
			}
			System.out.println("File Read!");
			workbook.close();
			
		

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

}