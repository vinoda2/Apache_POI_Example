package Demo_Excel;


	import java.io.File;
	import java.io.FileInputStream;
	import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class ReadData {
		public static void main(String[] args) {
			try {
				File file = new File("C:\\Users\\Windows\\Documents\\Data_List.xlsx"); 
				FileInputStream fis = new FileInputStream(file);

				HSSFWorkbook wb = new HSSFWorkbook(fis);
				HSSFSheet sheet = wb.getSheetAt(0);
				Iterator<Row> itr = sheet.iterator();
				while (itr.hasNext()) {
					Row row = itr.next();
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
