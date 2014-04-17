package techersjob;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
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
//		OutputStream out = new BufferedOutputStream(new FileOutputStream(outF));
//		XSSFWorkbook newWorkbok = new XSSFWorkbook();
		
		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(file);
		XSSFSheet sheet = workbook.getSheetAt(1);
		HATIVA hativaType = HATIVA.UPPER;
		double frontalSum = getFrontalSum(sheet,hativaType);
		double gmolSum = getGmolSum(sheet,hativaType);
		int teachersAge = getTeacherAge(sheet);
		
		
		System.out.println("frontal: "+frontalSum);
		System.out.println("gmol: "+gmolSum);
		System.out.println("age: "+teachersAge);
		//xlsxFile.createSheet("hello2");
		//XSSFSheet sheet = workbook.createSheet("hello");
//		Writer.createCell(workbook, sheet.createRow(0), 0, "100", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
//				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		
		Writer.write(workbook, "temp.xlsx");
		System.out.println("end");
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
		String date = cell.toString();
		DateFormat df = new SimpleDateFormat("dd-yyyy-mmm");
		try {
			Date birthDate = df.parse(date);
			double yearDiff = new Date().getYear() - birthDate.getYear();
			double monthsDiff = new Date().getMonth() - birthDate.getMonth();
			return (int)(yearDiff*12 + monthsDiff)/12;
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		return -1;
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
