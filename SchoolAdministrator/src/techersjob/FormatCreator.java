package techersjob;
import java.io.File;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import bagruyot.ExcellReaderAndWriter;
import bagruyot.Writer;

public class FormatCreator {

	public static boolean createWorkbook(int numOfSheets)
	{
		String pathname = "format.xlsx";
		File file = new File(pathname);
		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(file);
		workbook.setSheetName(0, "מורה " + 1);
		for (int i = 1; i < numOfSheets; i++) {
			workbook.cloneSheet(0);
			workbook.setSheetName(i, "מורה " + (i + 1));			
		}
		Writer.write(workbook, "result.xlsx");
		return true;
	}
	
}
