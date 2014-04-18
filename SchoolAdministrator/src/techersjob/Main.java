package techersjob;

import java.io.File;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import techersjob.TeacherAttributes.HATIVA;
import bagruyot.ExcellReaderAndWriter;
import bagruyot.Writer;

public class Main {

	

	public static void main(String[] args) throws IOException {
		System.out.println("start");
		String pathname = "temp.xlsx";
		File file = new File(pathname);

		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(file);
		XSSFSheet sheet = workbook.getSheetAt(1);
		HATIVA hativaType = HATIVA.UPPER;
		
		
		TeacherAttributes teacherAttributes = TeacherAttributes.extractDatafromSheet(sheet, hativaType);
		
		
		
		

		Writer.write(workbook, "temp.xlsx");
		System.out.println("end");
	}
	

	



	


	

}
