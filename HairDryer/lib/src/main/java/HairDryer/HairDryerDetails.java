package HairDryer;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class HairDryerDetails {
	
	public static void main(String[] args) {
		try {
			File file = new File("E:\\Excel_Data\\Hairdryer_Data.xlsx"); 
			FileInputStream input = new FileInputStream(file);

			HSSFWorkbook workbook = new HSSFWorkbook(input);
			HSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> it = sheet.iterator();
			while (it.hasNext()) {
				Row row = it.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					System.out.println(cell.toString());
				}
				System.out.println("");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
