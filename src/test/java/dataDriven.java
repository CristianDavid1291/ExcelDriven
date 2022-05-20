import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public static void main(String[] args) throws IOException {

		List<String> data = getData("Purchase");
		data.stream().filter(s->!s.equalsIgnoreCase("Purchase")).forEach(s->System.out.println(s));
		
	}

	public static List<String> getData(String testCase) throws IOException {

		List<String> a = new ArrayList<String>();
		// fileInputStream argument
		FileInputStream fis = new FileInputStream("F:\\CRISTIAN\\Eclipse\\ExcelDriven\\TestData.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		XSSFSheet sheet = workbook.getSheet("Hoja1");
		Iterator<Row> rows = sheet.rowIterator();
		Row firstRow = rows.next();
		Iterator<Cell> cells = firstRow.cellIterator();
		int k = 0;
		int column = 0;
		while (cells.hasNext()) {
			Cell value = cells.next();
			if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
				column = k;
			}
			k++;
		}
		while (rows.hasNext()) {
			Row row = rows.next();
			if (row.getCell(column).getStringCellValue().equalsIgnoreCase(testCase)) {
				Iterator<Cell> cv = row.cellIterator();
				while (cv.hasNext()) {
					
					Cell c = cv.next();
					if(c.getCellType() == CellType.STRING)
					{
						a.add(c.getStringCellValue());
					}else {
						a.add(String.valueOf(c.getNumericCellValue()));
					}
					
				}
			}
		}

		return a;

	}

}
