package techersjob;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import bagruyot.ExcellReaderAndWriter;
import bagruyot.Writer;

public class Main {

	enum HATIVA {UPPER, MEDIUM}
	private static final int TEACHER_AGE_ROW = 8;
	private static final int TEACHER_AGE_COLUMN = 8;


	public static void main(String[] args) throws IOException {
		System.out.println("start");
		String pathname = "temp.xlsx";
		File file = new File(pathname);

		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(file);
		XSSFSheet sheet = workbook.getSheetAt(1);
		HATIVA hativaType = HATIVA.UPPER;
		double frontalSum = getFrontalSum(sheet,hativaType);
		double gmolSum = getGmolSum(sheet,hativaType);
		int teachersAge = getTeacherAge(sheet);
		boolean isMothertoChileover14 = getIsMothertoChileover14(sheet);
		boolean perticipateinozTmura = getIsPerticipatesInosTmura(sheet);
		 
		System.out.println("frontal: "+frontalSum);
		System.out.println("gmol: "+gmolSum);
		System.out.println("age: "+teachersAge);
		System.out.println("mother to child over 14: "+isMothertoChileover14);
		System.out.println("participates in Oz Tmura: "+perticipateinozTmura);

		Writer.write(workbook, "temp.xlsx");
		System.out.println("end");
	}


	private static boolean getIsMothertoChileover14(XSSFSheet sheet) {
		XSSFRow row = sheet.getRow(TEACHER_AGE_ROW+1);
		XSSFCell cell = row.getCell(TEACHER_AGE_COLUMN);
		return (cell.getStringCellValue().equals("λο"))? true:false;
	}


	private static boolean getIsPerticipatesInosTmura(XSSFSheet sheet) {
		XSSFRow row = sheet.getRow(TEACHER_AGE_ROW+2);
		XSSFCell cell = row.getCell(TEACHER_AGE_COLUMN);
		return (cell.getStringCellValue().equals("λο"))? true:false;
	}


	//	private static double calculatePartani(int hours, HATIVA hativaType){
	//		
	//	}

	private static int getHativaRow(HATIVA hativaType){
		int row = -1;
		switch(hativaType){
		case UPPER:
			row = 25;
			break;
		case MEDIUM:
			row = 24;
		}	
		return row;
	}

	private static int getTeacherAge(XSSFSheet sheet){
		XSSFRow row = sheet.getRow(TEACHER_AGE_ROW);
		XSSFCell cell = row.getCell(TEACHER_AGE_COLUMN);
		Date dateCellValue = cell.getDateCellValue();
		Calendar birthDayCal = Calendar.getInstance();
		birthDayCal.setTime(dateCellValue);
		Calendar currentCal = Calendar.getInstance();
		currentCal.setTime(new Date());

		int yearOld = currentCal.get(Calendar.YEAR) - birthDayCal.get(Calendar.YEAR);
		if(currentCal.get(Calendar.MONTH) < birthDayCal.get(Calendar.MONTH)){
			yearOld--;
		}
		return yearOld;
	}



	/**
	 * getting the frontal sum
	 * 
	 * @param sheet
	 * @param hativaType
	 * @return
	 */
	private static double getFrontalSum(XSSFSheet sheet, HATIVA hativaType){
		int rowNum = getHativaRow(hativaType);
		XSSFRow row = sheet.getRow(rowNum);
		XSSFCell cell = row.getCell(4);
		double value = cell.getNumericCellValue();
		return value;
	}


	/**
	 * getting the gmol sum
	 * 
	 * @param sheet
	 * @param hativaType
	 * @return
	 */
	private static double getGmolSum(XSSFSheet sheet, HATIVA hativaType){
		int rowNum = getHativaRow(hativaType);
		XSSFRow row = sheet.getRow(rowNum);
		XSSFCell cell = row.getCell(7);
		double value = cell.getNumericCellValue();
		return value;
	}

}
