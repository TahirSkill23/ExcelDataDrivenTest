import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream("C://Users//213642//Documents//TestDemoDataforExceldrivenTest.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetcount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetcount; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.rowIterator();// Sheet is a collection of rows
				Row firstrow = rows.next();
				Iterator<Cell> cells = firstrow.cellIterator();// row is a collection of cell
				int k=0;
				int column=0;
				while (cells.hasNext()) {
					Cell cellvalue = cells.next();
					if (cellvalue.getStringCellValue().equalsIgnoreCase("Testcases")) {
                        column=k;
					}
					k++;
				}
				System.out.println(column);
				while(rows.hasNext()) {
				Row r=rows.next();
		if(	r.getCell(column).getStringCellValue().equalsIgnoreCase("update profile")) {
		Iterator<Cell>	cv=r.cellIterator();
		while(cv.hasNext()) {
			System.out.println(cv.next().getStringCellValue());
		}
		}
			   
				}
			}
		}
	}

}
