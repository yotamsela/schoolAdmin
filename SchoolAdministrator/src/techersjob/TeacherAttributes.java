package techersjob;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class TeacherAttributes {
	
	enum HATIVA {UPPER, MEDIUM}
	private static final int TEACHER_AGE_ROW = 8;
	private static final int TEACHER_AGE_COLUMN = 8;

	
	private double frontalHours;
	private double gmol;
	private int teacherAge;
	private boolean ozTmura;
	private boolean childOver14;

	private TeacherAttributes(double frontalHours, double gmol, int teacherAge,
			boolean ozTmura, boolean childOver14) {
		this.frontalHours = frontalHours;
		this.gmol = gmol;
		this.teacherAge = teacherAge;
		this.ozTmura = ozTmura;
		this.childOver14 = childOver14;
	}
	
	public double getFrontalHours() {
		return frontalHours;
	}

	public double getGmol() {
		return gmol;
	}

	public int getTeacherAge() {
		return teacherAge;
	}

	public boolean isOzTmura() {
		return ozTmura;
	}

	public boolean isChildOver14() {
		return childOver14;
	}

	public static TeacherAttributes extractDatafromSheet(XSSFSheet sheet,HATIVA hativaType){
		double frontalHours = getFrontalSum(sheet,hativaType);
		double gmol = getGmolSum(sheet,hativaType);
		int teacherAge = getTeacherAge(sheet);
		boolean childOver14 = getIsMothertoChileover14(sheet);
		boolean ozTmura = getIsPerticipatesInosTmura(sheet);
		 
		System.out.println("frontal: "+frontalHours);
		System.out.println("gmol: "+gmol);
		System.out.println("age: "+teacherAge);
		System.out.println("mother to child over 14: "+childOver14);
		System.out.println("participates in Oz Tmura: "+ozTmura);
		
		return new TeacherAttributes(frontalHours, gmol, teacherAge, ozTmura, childOver14);
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
