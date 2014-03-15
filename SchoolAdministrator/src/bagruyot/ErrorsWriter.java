package bagruyot;
import java.util.Set;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ErrorsWriter {

	public static void write(XSSFSheet dataSheet, Set<Integer> unrecognizedQuestionnaires, Set<Integer> unrecognizedGroups, String name) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet;
		if(unrecognizedQuestionnaires.size() > 0){
			sheet = workbook.createSheet("שאלונים לא מוכרים");
			fillSheet(sheet, dataSheet, unrecognizedQuestionnaires);
		}
		if(unrecognizedGroups.size() > 0){
			sheet = workbook.createSheet("קבוצות לא מוכרות");
			fillSheet(sheet, dataSheet, unrecognizedQuestionnaires);
		}
		Writer.write(workbook, name + " - שגיאות.xlsx");
	}

	private static void fillSheet(XSSFSheet sheet, XSSFSheet dataSheet, Set<Integer> unrecognizedQuestionnaires) {
		sheet.setRightToLeft(true);
		int rowNum = 0;
		XSSFRow dataRow, row = null;
		for (Integer i: unrecognizedQuestionnaires) {
			row = sheet.createRow(rowNum++);
			dataRow = dataSheet.getRow(i);
			ExcellReaderAndWriter.copyRow(dataRow, row);
		}
		int numOfColumns = row.getLastCellNum();
		String end = "";
		for (int i = 0; i < (numOfColumns / 26); i++) {
			end = end + 'A';
		}
		end = end + (char)('A' + (numOfColumns % 26));
		sheet.setAutoFilter(CellRangeAddress.valueOf("A1:" + end + (rowNum - 1)));
	}

}
