package techersjob;

import java.io.File;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import bagruyot.ExcellReaderAndWriter;

import techersjob.TabeleSeperator.TableType;

public class TeacherHoursOz implements ITeacherTable{
	
	private Map<TabeleSeperator, Map<Double, TeacherHours>> map;
	private TableType type = TableType.OZ;

	@Override
	public void initTableFromfile(
			String fileName) {
		
	}

	@Override
	public TeacherHours getTeacherTable(int age, double teachingHoursKey) {
		TabeleSeperator range = TabeleSeperator.getRange(age,type);
		if(range == null){
			return null;
		}
		Map<Double, TeacherHours> innerMap = map.get(range);
		return innerMap.get(teachingHoursKey);
	}

	@Override
	public void extractFromExelToTxtFile(String fileName) {
		File file = new File(fileName);

		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		for (int i = 0; i < 10; i++) {
			XSSFRow row = sheet.getRow(i);
			for (int j = 0; j < 10; j++) {
				XSSFCell cell = row.getCell(j);
				
			}
			
		}
		
		
		
		
	}

}
