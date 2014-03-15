package bagruyot;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class StudentsWriter {

	public static final String FIRST_PAGE_NAME = "סיכום";

	public static void write(List<Subject> subjects, List<Student> students, String name) {
		if(students.size() == 0){
			//TODO print message
			return;
		}
		XSSFWorkbook workbook = new XSSFWorkbook();
		printStatistics(subjects, students, workbook);
		for (Student student: students) {
			printStudent(student, workbook);			
		}
		Writer.write(workbook, name + " - ממוצעי תלמידים.xlsx");
	}

	private static void printStatistics(List<Subject> subjects, List<Student> students, XSSFWorkbook workbook) {
		XSSFSheet sheet = workbook.createSheet(FIRST_PAGE_NAME);
		sheet.setRightToLeft(true);
		int numOfColumns = printStatisticsHeader(workbook, sheet, subjects);
		List<List<Double>> averages = printStatisticsBody(workbook, sheet, subjects, students, numOfColumns);
		printStatisticsSummary(workbook, sheet, subjects, students, averages);
		sheet.createFreezePane(0, 1, 0, 1);
	}

	private static int printStatisticsHeader(XSSFWorkbook workbook, XSSFSheet sheet, List<Subject> subjects) {
		XSSFRow row = sheet.createRow(0);
		int colNum = 0;
		Writer.createCell(workbook, row, colNum++, "שם התלמיד", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, false, (short)12, false, false);
		for (Subject subject : subjects) {
			Writer.createCell(workbook, row, colNum++, subject.getName() + " - " + subject.getNumOfUnits() + " יח\"ל", 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, 
					(short)12, false, true);
		}
		Writer.createCell(workbook, row, colNum++, "ממוצע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)12, false, true);
		Writer.createCell(workbook, row, colNum++, "ממוצע משוקלל", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)12, false, true);
		Writer.createCell(workbook, row, colNum++, "מספר יח\"ל", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)12, false, true);
		return colNum;
	}

	private static List<List<Double>> printStatisticsBody(XSSFWorkbook workbook, XSSFSheet sheet, List<Subject> subjects, List<Student> students,
			int numOfColumns) {
		int colNum, rowNum = 1;
		Subject studentSubject;
		XSSFRow row = null;
		List<Subject> studentsSubjects;
		List<List<Double>> averages = new ArrayList<List<Double>>();
		XSSFCellStyle style;
		averages.add(new ArrayList<Double>());
		averages.add(new ArrayList<Double>());
		averages.add(new ArrayList<Double>());
		averages.add(new ArrayList<Double>());
		for (int i = 0; i <= subjects.size(); i++) {
			averages.get(0).add(0.0);
			averages.get(1).add(0.0);
			averages.get(2).add(0.0);
			averages.get(3).add(0.0);			
		}
		averages.get(3).add(0.0);
		averages.get(3).add(0.0);
		for (Student student : students) {
			//TODO ask?
//			if(student.getSchoolGrade() >= SchoolClass.LAST_GRADE){
				colNum = 0;
				row = sheet.createRow(rowNum++);
				Writer.createCell(workbook, row, colNum++, student.getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, (short)10, IndexedColors.BLUE.getIndex(), 
						IndexedColors.WHITE.getIndex(), true, false, false, true);
				for (Subject subject : subjects) {
					studentsSubjects = student.getRelevantSubjects();
					studentSubject = null;
					for (Subject s: studentsSubjects) {
						if(s.equals(subject) && s.getNumOfUnits() == subject.getNumOfUnits()){
							studentSubject = s;
						}
					}
					if(studentSubject != null){
						averages.get(0).set(colNum - 1, averages.get(0).get(colNum - 1) + 1.0);
						if(studentSubject.isCompleated()){
							averages.get(1).set(colNum - 1, averages.get(1).get(colNum - 1) + 1.0);
							if(studentSubject.hasPassed()){
								averages.get(2).set(colNum - 1, averages.get(2).get(colNum - 1) + 1.0);
								averages.get(3).set(colNum - 1, averages.get(3).get(colNum - 1) + studentSubject.getFinalGrade());
								Writer.createCell(workbook, row, colNum++, studentSubject.getFinalGrade(), XSSFCellStyle.BORDER_THIN, 
										XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, (short)11, 
										IndexedColors.BLACK.getIndex(), IndexedColors.WHITE.getIndex(), false, false, false, false);
							}
							else{
								if(studentSubject.isMandatory()){
									Writer.createCell(workbook, row, colNum++, studentSubject.getFinalGrade(), XSSFCellStyle.BORDER_THIN, 
											XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, (short)11, 
											IndexedColors.BLACK.getIndex(), IndexedColors.RED.getIndex(), false, false, false, false);									
								}
								else{
									Writer.createCell(workbook, row, colNum++, studentSubject.getFinalGrade(), XSSFCellStyle.BORDER_THIN, 
											XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, (short)11, 
											IndexedColors.BLACK.getIndex(), IndexedColors.YELLOW.getIndex(), false, false, false, false);
								}
							}
						}
						else{
							if(studentSubject.isMandatory()){
								Writer.createCell(workbook, row, colNum++, (int)studentSubject.getPercentage() + "%", XSSFCellStyle.BORDER_THIN, 
										XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, (short)11, 
										IndexedColors.BLACK.getIndex(), IndexedColors.RED.getIndex(), false, false, false, false);
							}
							else{
								Writer.createCell(workbook, row, colNum++, "", XSSFCellStyle.BORDER_THIN, 
										XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, (short)11, 
										IndexedColors.BLACK.getIndex(), IndexedColors.WHITE.getIndex(), false, false, false, false);								
							}
						}
					}
					else{
						Writer.createCell(workbook, row, colNum++, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
								XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)10, true, false);
					}
				}
				averages.get(0).set(averages.get(0).size() - 1, averages.get(0).get(averages.get(0).size() - 1) + 1);
				if(student.isEntitled()){
					Writer.createCell(workbook, row, colNum++, student.getFinalGrade(), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, (short)10, 
							IndexedColors.BLACK.getIndex(), IndexedColors.WHITE.getIndex(), true, false, false, false, 2);
					Writer.createCell(workbook, row, colNum++, student.getWeightedGrade(), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, (short)10, 
							IndexedColors.BLACK.getIndex(), IndexedColors.WHITE.getIndex(), true, false, false, false, 2);
					Writer.createCell(workbook, row, colNum++, student.getNumOfUnits(), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, (short)10, 
							IndexedColors.BLACK.getIndex(), IndexedColors.WHITE.getIndex(), true, false, false, false, 2);
					averages.get(1).set(averages.get(1).size() - 1, averages.get(1).get(averages.get(1).size() - 1) + 1);
					averages.get(3).set(averages.get(3).size() - 3, averages.get(3).get(averages.get(3).size() - 3) + student.getFinalGrade());
					averages.get(3).set(averages.get(3).size() - 2, averages.get(3).get(averages.get(3).size() - 2) + student.getWeightedGrade());
					averages.get(3).set(averages.get(3).size() - 1, averages.get(3).get(averages.get(3).size() - 1) + student.getNumOfUnits());
					//TODO set the color of the student's average
				}
				else{
					Writer.createCell(workbook, row, colNum++, student.getFinalGrade(), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, (short)10, 
							IndexedColors.BLACK.getIndex(), IndexedColors.RED.getIndex(), true, false, false, false, 2);
					Writer.createCell(workbook, row, colNum++, student.getWeightedGrade(), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, (short)10, 
							IndexedColors.BLACK.getIndex(), IndexedColors.RED.getIndex(), true, false, false, false, 2);
					Writer.createCell(workbook, row, colNum++, student.getNumOfUnits(), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, (short)10, 
							IndexedColors.BLACK.getIndex(), IndexedColors.RED.getIndex(), true, false, false, false, 2);					
				}
//			}
		}
		for (int i = 0; i < averages.get(1).size() - 1; i++) {
			if(averages.get(1).get(i) == 0){				
				averages.get(3).set(i, 0.0);
			}
			else{
				averages.get(3).set(i, averages.get(3).get(i) / averages.get(2).get(i));
			}
		}
		averages.get(3).set(averages.get(1).size() - 1, averages.get(3).get(averages.get(1).size() - 1) / averages.get(1).get(averages.get(1).size() - 1));
		averages.get(3).set(averages.get(1).size(), averages.get(3).get(averages.get(1).size()) / averages.get(1).get(averages.get(1).size() - 1));
		averages.get(3).set(averages.get(1).size() + 1, averages.get(3).get(averages.get(1).size() + 1) / averages.get(1).get(averages.get(1).size() - 1));
		style = row.getCell(0).getCellStyle();
		Writer.changeCellStyle(workbook, row.getCell(0), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), style.getBorderRight(),
			0);
		for (int j = 1; j < numOfColumns - 3; j++) {
			style = row.getCell(j).getCellStyle();
		Writer.changeCellStyle(workbook, row.getCell(j), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), 
			style.getBorderRight(), 0);
		}
		style = row.getCell(numOfColumns - 3).getCellStyle();
		Writer.changeCellStyle(workbook, row.getCell(numOfColumns - 3), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), 
			style.getBorderRight(), 2);
		style = row.getCell(numOfColumns - 2).getCellStyle();
		Writer.changeCellStyle(workbook, row.getCell(numOfColumns - 2), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), 
			style.getBorderRight(), 2);
		style = row.getCell(numOfColumns - 1).getCellStyle();
		Writer.changeCellStyle(workbook, row.getCell(numOfColumns - 1), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), 
			style.getBorderRight(), 0);
		String end = "";
		for (int i = 0; i < ((numOfColumns - 1) / 26); i++) {
			end = end + 'A';
		}
		end = end + (char)('A' + ((numOfColumns - 1) % 26));
		sheet.setAutoFilter(CellRangeAddress.valueOf("A1:" + end + (row.getRowNum() - 1)));
		return averages;
	}

	private static void printStatisticsSummary(XSSFWorkbook workbook, XSSFSheet sheet, List<Subject> subjects, List<Student> students, 
			List<List<Double>> averages) {
		int numOfSubjects = averages.get(0).size() - 1, rowNum = (int) (averages.get(0).get(numOfSubjects) + 1);
		XSSFRow firstRow = sheet.createRow(rowNum++), secondRow = null, thirdRow = null, forthRow = null;
		XSSFCellStyle style;
		Writer.createCell(workbook, firstRow, 0, "סיכום", XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);
		for (int i = 1; i <= numOfSubjects + 3; i++) {
			Writer.createCell(workbook, firstRow, i, "", XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);			
		}
		sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, numOfSubjects + 3));
		firstRow = sheet.createRow(rowNum++);
		secondRow = sheet.createRow(rowNum++);
		thirdRow = sheet.createRow(rowNum++);
		forthRow = sheet.createRow(rowNum++);		
		Writer.createCell(workbook, firstRow, 0, "מספר ניגשים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, secondRow, 0, "מספר מסיימים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, thirdRow, 0, "מספר עוברים", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, forthRow, 0, "ממוצע ציונים של העוברים", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		sheet.setColumnWidth(0, 6300);
		for (int j = 0; j <= numOfSubjects; j++) {
			Writer.createCell(workbook, firstRow, j + 1, (int)(double)averages.get(0).get(j), XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)10);
			Writer.createCell(workbook, secondRow, j + 1, (int)(double)averages.get(1).get(j), XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)10);
			Writer.createCell(workbook, thirdRow, j + 1, (int)(double)averages.get(2).get(j), XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)10);
			Writer.createCell(workbook, forthRow, j + 1, averages.get(3).get(j), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)10, 1);
			sheet.setColumnWidth(j, 1000);
		}
		sheet.setColumnWidth(0, 5000);
		style = firstRow.getCell(1).getCellStyle();
		Writer.changeCellStyle(workbook, firstRow.getCell(1), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
			style.getBorderRight(), 0);
		style = secondRow.getCell(1).getCellStyle();
		Writer.changeCellStyle(workbook, secondRow.getCell(1), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
			style.getBorderRight(), 0);
		style = thirdRow.getCell(1).getCellStyle();
		Writer.changeCellStyle(workbook, thirdRow.getCell(1), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
			style.getBorderRight(), 0);
		style = forthRow.getCell(1).getCellStyle();
		Writer.changeCellStyle(workbook, forthRow.getCell(1), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
			style.getBorderRight(), 1);
		style = firstRow.getCell(numOfSubjects).getCellStyle();
		Writer.changeCellStyle(workbook, firstRow.getCell(numOfSubjects), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
			XSSFCellStyle.BORDER_MEDIUM, 0);
		style = secondRow.getCell(numOfSubjects).getCellStyle();
		Writer.changeCellStyle(workbook, secondRow.getCell(numOfSubjects), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
			XSSFCellStyle.BORDER_MEDIUM, 0);
		style = thirdRow.getCell(numOfSubjects).getCellStyle();
		Writer.changeCellStyle(workbook, thirdRow.getCell(numOfSubjects), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
			XSSFCellStyle.BORDER_MEDIUM, 0);
		style = forthRow.getCell(numOfSubjects).getCellStyle();
		Writer.changeCellStyle(workbook, forthRow.getCell(numOfSubjects), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
			XSSFCellStyle.BORDER_MEDIUM, 1);
		Writer.createCell(workbook, firstRow, numOfSubjects + 1, (int)(double)averages.get(0).get(numOfSubjects), XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10);
		Writer.createCell(workbook, secondRow, numOfSubjects + 1, (int)(double)averages.get(1).get(numOfSubjects), XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10);
		Writer.createCell(workbook, thirdRow, numOfSubjects + 1, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, forthRow, numOfSubjects + 1, averages.get(3).get(numOfSubjects), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, 2);
		Writer.createCell(workbook, firstRow, numOfSubjects + 2, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, secondRow, numOfSubjects + 2, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, thirdRow, numOfSubjects + 2, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, forthRow, numOfSubjects + 2, averages.get(3).get(numOfSubjects + 1), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, 2);
		Writer.createCell(workbook, firstRow, numOfSubjects + 3, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, secondRow, numOfSubjects + 3, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, thirdRow, numOfSubjects + 3, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)10, false, false);
		Writer.createCell(workbook, forthRow, numOfSubjects + 3, averages.get(3).get(numOfSubjects + 2), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, 2);
		sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), numOfSubjects + 1, numOfSubjects + 3));
		sheet.addMergedRegion(new CellRangeAddress(secondRow.getRowNum(), thirdRow.getRowNum(), numOfSubjects + 1, numOfSubjects + 3));
		sheet.setColumnWidth(numOfSubjects + 1, 1900);
		sheet.setColumnWidth(numOfSubjects + 2, 1900);
		sheet.setColumnWidth(numOfSubjects + 3, 1200);
		firstRow = sheet.createRow(rowNum++);
		Writer.createCell(workbook, firstRow, numOfSubjects + 1, "אחוז זכאים לבגרות:", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, true, (short)10, false, false);
		Writer.createCell(workbook, firstRow, numOfSubjects + 2, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)10, false, false);
		sheet.addMergedRegion(new CellRangeAddress(firstRow.getRowNum(), firstRow.getRowNum(), numOfSubjects + 1, numOfSubjects + 2));
		Writer.createCell(workbook, firstRow, numOfSubjects + 3, averages.get(1).get(numOfSubjects) / averages.get(0).get(numOfSubjects) * 100, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)10, 2);
	}

	private static void printStudent(Student student, XSSFWorkbook workbook) {
		String name = student.getName();
		if(name.endsWith("\'")){
			name = name.substring(0, name.length() - 1);
		}
		int numOfExams = 0;
		for (Subject subject : student.getRelevantSubjects()) {
			numOfExams = Math.max(numOfExams, subject.getRelevantQuestionaires().keySet().size());
		}
		numOfExams--;
		XSSFSheet sheet = workbook.createSheet(name);
		sheet.setRightToLeft(true);
		CreationHelper helper = workbook.getCreationHelper();
		createHyperLink(workbook, sheet, helper, student.getName());
		printStudentSubjectsHeader(workbook, sheet, student, helper);
		int rowNum = printStudentSubjects(workbook, sheet, student, numOfExams);
		printStudentSummary(workbook, sheet, student, rowNum);
		printStudentExamsHeader(workbook, sheet, student, helper, numOfExams);
		printStudentExams(workbook, sheet, student, numOfExams);
	}

	private static void createHyperLink(XSSFWorkbook workbook, XSSFSheet sheet, CreationHelper helper, String name) {
		XSSFSheet firstSheet = workbook.getSheetAt(0);
		Hyperlink link;
		for (Row row : firstSheet) {
			if(row.getRowNum() > 0 && row.getCell(0) != null && row.getCell(0).toString().equals(name)){
				link = helper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link.setAddress("'" + name + "'!A1");
				row.getCell(0).setHyperlink(link);
			}
		}
	}

	private static void printStudentSubjectsHeader(XSSFWorkbook workbook, XSSFSheet sheet, Student student, CreationHelper helper){
		int rownum = 0;
		Hyperlink link;
		XSSFRow row = sheet.createRow(rownum++);
		Writer.createCell(workbook, row, 1, "פרטים אישיים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)18, false, false);
		Writer.createCell(workbook, row, 2, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		Writer.createCell(workbook, row, 3, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		Writer.createCell(workbook, row, 4, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		Writer.createCell(workbook, row, 5, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		Writer.createCell(workbook, row, 6, "בחזרה לדף הראשי", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, (short)12, IndexedColors.BLUE.getIndex(), IndexedColors.WHITE.getIndex(), true, true, false, true);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 5));
		row = sheet.createRow(rownum++);
		Writer.createCell(workbook, row, 1, "שם התלמיד:", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, 2, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, 3, "ת.ז.", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, 4, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, 5, "כיתה:", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);		
		Writer.createCell(workbook, row, 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)16, true, false);
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 4));
		row = sheet.createRow(rownum++);
		Writer.createCell(workbook, row, 1, student.getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
		Writer.createCell(workbook, row, 2, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, 3, student.getIdNumber(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
		Writer.createCell(workbook, row, 4, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, 5, student.getSchoolClass().getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, false, (short)11, false, false);
		Writer.createCell(workbook, row, 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)16, true, false);
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 2));
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 4));
		sheet.addMergedRegion(new CellRangeAddress(0, 2, 6, 6));
		link = helper.createHyperlink(Hyperlink.LINK_DOCUMENT);
		link.setAddress("'" + FIRST_PAGE_NAME + "'!A1");
		sheet.getRow(0).getCell(6).setHyperlink(link);
	}

	private static int printStudentSubjects(XSSFWorkbook workbook, XSSFSheet sheet, Student student, int numOfExams) {
		int rowNum = 4, colNum;
		Map<Questionnaire, Double> relevantQuestionaires;
		XSSFRow firstRow = sheet.createRow(rowNum++), secondRow = null, thirdRow = null;
		XSSFCellStyle style;
		Writer.createCell(workbook, firstRow, 0, "מקצועות", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)18, false, false);
		for (int i = 1; i <= numOfExams + 3; i++) {
			Writer.createCell(workbook, firstRow, i, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)18, false, false);
		}
		sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, numOfExams + 3));
		firstRow = sheet.createRow(rowNum++);
		Writer.createCell(workbook, firstRow, 0, "מקצוע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, firstRow, 1, "מספר יח\"ל", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, firstRow, 2, "ציון סופי", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
		for (int i = 1; i <= numOfExams; i++) {
			Writer.createCell(workbook, firstRow, 2 + i, "מבחן - " + i, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);			
		}
		style = firstRow.getCell(3).getCellStyle();
		Writer.changeCellStyle(workbook, firstRow.getCell(3), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
			style.getBorderRight(), 0);
		style = firstRow.getCell(numOfExams + 1).getCellStyle();
		Writer.changeCellStyle(workbook, firstRow.getCell(numOfExams + 1), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
			XSSFCellStyle.BORDER_MEDIUM, 0);
		Writer.createCell(workbook, firstRow, numOfExams + 3, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
		for (Subject subject : student.getRelevantSubjects()) {
			firstRow = sheet.createRow(rowNum++);
			secondRow = sheet.createRow(rowNum++);
			thirdRow = sheet.createRow(rowNum++);
			Writer.createCell(workbook, firstRow, 0, subject.getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
			Writer.createCell(workbook, secondRow, 0, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
			Writer.createCell(workbook, thirdRow, 0, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_DOUBLE, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 3, rowNum - 1, 0, 0));
			Writer.createCell(workbook, firstRow, 1, subject.getNumOfUnits(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11);
			Writer.createCell(workbook, secondRow, 1, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
			Writer.createCell(workbook, thirdRow, 1, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_DOUBLE, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 3, rowNum - 1, 1, 1));
			if(subject.isCompleated()){
				Writer.createCell(workbook, firstRow, 2, subject.getFinalGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11);				
			}
			else{
				Writer.createCell(workbook, firstRow, 2, "לא הושלם", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);			
			}
			Writer.createCell(workbook, secondRow, 2, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
			Writer.createCell(workbook, thirdRow, 2, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_DOUBLE, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
			sheet.addMergedRegion(new CellRangeAddress(rowNum - 3, rowNum - 1, 2, 2));
			colNum = 3;
			relevantQuestionaires = subject.getRelevantQuestionaires();
			for (Iterator<Questionnaire> iterator = relevantQuestionaires.keySet().iterator(); iterator.hasNext();) {
				Questionnaire questionnaire = iterator.next();
				for (Questionnaire q : subject.getQuestionnaires()) {
					if(q.equals(questionnaire)){
						questionnaire = q;
						break;
					}
				}
				if(questionnaire != null && questionnaire.getBestExam() != null){
					Writer.createCell(workbook, firstRow, colNum, questionnaire.getNumber(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
					Writer.createCell(workbook, secondRow, colNum, relevantQuestionaires.get(questionnaire), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, 2);
					Writer.createCell(workbook, thirdRow, colNum++, questionnaire.getBestExam().getFinalGrade(), XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_DOUBLE, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
				}
			}
			while(colNum < numOfExams + 3) {
				Writer.createCell(workbook, firstRow, colNum, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
				Writer.createCell(workbook, secondRow, colNum, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
				Writer.createCell(workbook, thirdRow, colNum++, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_DOUBLE, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
			}
			style = firstRow.getCell(3).getCellStyle();
			Writer.changeCellStyle(workbook, firstRow.getCell(3), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
				style.getBorderRight(), 0);
			style = firstRow.getCell(numOfExams + 1).getCellStyle();
			Writer.changeCellStyle(workbook, firstRow.getCell(numOfExams + 1), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
				XSSFCellStyle.BORDER_MEDIUM, 0);
			style = secondRow.getCell(3).getCellStyle();
			Writer.changeCellStyle(workbook, secondRow.getCell(3), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
				style.getBorderRight(), 2);
			style = secondRow.getCell(numOfExams + 1).getCellStyle();
			Writer.changeCellStyle(workbook, secondRow.getCell(numOfExams + 1), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
				XSSFCellStyle.BORDER_MEDIUM, 2);
			style = thirdRow.getCell(3).getCellStyle();
			Writer.changeCellStyle(workbook, thirdRow.getCell(3), style.getBorderTop(), style.getBorderBottom(), XSSFCellStyle.BORDER_MEDIUM, 
				style.getBorderRight(), 0);
			style = thirdRow.getCell(numOfExams + 1).getCellStyle();
			Writer.changeCellStyle(workbook, thirdRow.getCell(numOfExams + 1), style.getBorderTop(), style.getBorderBottom(), style.getBorderLeft(), 
				XSSFCellStyle.BORDER_MEDIUM, 0);
			Writer.createCell(workbook, firstRow, numOfExams + 3, "מספר שאלון", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, false, (short)11, false, false);			
			Writer.createCell(workbook, secondRow, numOfExams + 3, "משקל יחסי", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, false, (short)11, false, false);
			Writer.createCell(workbook, thirdRow, numOfExams + 3, "ציון סופי", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_DOUBLE, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, false, (short)11, false, false);
		}
		for (int i = 0; i < numOfExams + 4; i++) {
			style = thirdRow.getCell(i).getCellStyle();
			Writer.changeCellStyle(workbook, thirdRow.getCell(i), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), 
				style.getBorderRight(), 0);
		}
		sheet.setColumnWidth(0, 2050);
		sheet.setColumnWidth(1, 2200);
		sheet.setColumnWidth(2, 1850);
		for (int i = 3; i < numOfExams + 3; i++) {
			sheet.setColumnWidth(i, 1900);
		}
		sheet.setColumnWidth(numOfExams + 3, 2400);
		return rowNum;
	}

	private static void printStudentSummary(XSSFWorkbook workbook, XSSFSheet sheet, Student student, int rowNum) {		
		XSSFRow row = sheet.createRow(rowNum++);
		Writer.createCell(workbook, row, 0, "סה\"כ:", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)14, false, false);
		Writer.createCell(workbook, row, 1, student.getNumOfUnits(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)14);
		Writer.createCell(workbook, row, 2, student.getFinalGrade(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)14, 2);
		row = sheet.createRow(rowNum++);
		Writer.createCell(workbook, row, 0, ((student.isEntitled())? "": "לא ") + "זכאי/ת לתעודת בגרות", XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)12, false, false);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)14, false, false);
		Writer.createCell(workbook, row, 2, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)14, false, false);
		sheet.addMergedRegion(new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 2));
		row = sheet.createRow(rowNum++);
		Writer.createCell(workbook, row, 0, "ציון אוניברסיטאי משוקלל:", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)13, true, false);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, true, (short)14, false, false);
		Writer.createCell(workbook, row, 2, student.getWeightedGrade(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)12, 2);
		row = sheet.createRow(rowNum++);
		Writer.createCell(workbook, row, 0, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)14, true, false);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, true, (short)14, false, false);
		Writer.createCell(workbook, row, 2, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)14, false, false);
		row = sheet.createRow(rowNum++);
		Writer.createCell(workbook, row, 0, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)14, true, false);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, true, (short)14, false, false);
		sheet.addMergedRegion(new CellRangeAddress(rowNum - 3, rowNum - 1, 0, 1));
		Writer.createCell(workbook, row, 2, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)14, false, false);
		sheet.addMergedRegion(new CellRangeAddress(rowNum - 3, rowNum - 1, 2, 2));
	}

	private static void printStudentExamsHeader(XSSFWorkbook workbook, XSSFSheet sheet, Student student, CreationHelper helper, int numOfExams){
		int rownum = 0;
		XSSFRow row = sheet.getRow(rownum++);
		Writer.createCell(workbook, row, numOfExams + 6, "פרטים אישיים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)18, false, false);
		Writer.createCell(workbook, row, numOfExams + 7, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		Writer.createCell(workbook, row, numOfExams + 8, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		Writer.createCell(workbook, row, numOfExams + 9, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, numOfExams + 6, numOfExams + 9));
		row = sheet.getRow(rownum++);
		Writer.createCell(workbook, row, numOfExams + 6, "שם התלמיד:", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, numOfExams + 7, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, numOfExams + 8, "ת.ז.", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, numOfExams + 9, "כיתה:", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);		
		sheet.addMergedRegion(new CellRangeAddress(1, 1, numOfExams + 6, numOfExams + 7));
		row = sheet.getRow(rownum++);
		Writer.createCell(workbook, row, numOfExams + 6, student.getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
		Writer.createCell(workbook, row, numOfExams + 7, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, numOfExams + 8, student.getIdNumber(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
		Writer.createCell(workbook, row, numOfExams + 9, student.getSchoolClass().getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, false, (short)11, false, false);
		sheet.addMergedRegion(new CellRangeAddress(2, 2, numOfExams + 6, numOfExams + 7));
	}

	private static void printStudentExams(XSSFWorkbook workbook, XSSFSheet sheet, Student student, int numOfExams) {
		int rowNum = 4, colNum = numOfExams + 5;
		XSSFRow row = sheet.getRow(rowNum++);
		XSSFCellStyle style;
		Writer.createCell(workbook, row, colNum++, "מבחנים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)18, false, false);
		for (int i = 0; i < 5; i++) {
			Writer.createCell(workbook, row, colNum++, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)18, false, false);
		}
		sheet.addMergedRegion(new CellRangeAddress(4, 4, numOfExams + 5, numOfExams + 10));
		row = sheet.getRow(rowNum++);
		colNum = numOfExams + 5;
		Writer.createCell(workbook, row, colNum++, "מקצוע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, colNum++, "מס\' שאלון", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, colNum++, "תאריך", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, colNum++, "ציון מגן", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, colNum++, "ציון מבחן", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, row, colNum++, "ציון סופי", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
		colNum = numOfExams + 5;
		for (Subject subject : student.getRelevantSubjects()) {
			for (Questionnaire questionnaire : subject.getQuestionnaires()) {
				printExam(workbook, sheet, subject, questionnaire, questionnaire.getBestExam(), rowNum++, colNum);
				for (Exam exam : questionnaire.getExams()) {
					if(exam != questionnaire.getBestExam()){
						printExam(workbook, sheet, subject, questionnaire, exam, rowNum++, colNum);
					}
				}
			}
		}
		colNum = numOfExams + 5;
		for (int i = 0; i <= 5; i++) {
			style = sheet.getRow(rowNum - 1).getCell(colNum).getCellStyle();
			Writer.changeCellStyle(workbook, sheet.getRow(rowNum - 1).getCell(colNum++), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), 
				style.getBorderRight(), 0);
		}
		sheet.setColumnWidth(numOfExams + 5, 3000);
		sheet.setColumnWidth(numOfExams + 7, 2800);
		sheet.setColumnWidth(numOfExams + 8, 2800);
	}

	private static void printExam(XSSFWorkbook workbook, XSSFSheet sheet, Subject subject, Questionnaire questionnaire, Exam exam, int rowNum, 
			int colNum){
		XSSFRow row;
		if(sheet.getRow(rowNum) == null){
			row = sheet.createRow(rowNum);
		}
		else{
			row = sheet.getRow(rowNum);
		}
		Writer.createCell(workbook, row, colNum++, subject.getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
		Writer.createCell(workbook, row, colNum++, questionnaire.getNumber(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
		Writer.createCell(workbook, row, colNum++, exam.getDate(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
		Writer.createCell(workbook, row, colNum++, exam.getMagenGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
		Writer.createCell(workbook, row, colNum++, exam.getExamGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
		Writer.createCell(workbook, row, colNum++, exam.getFinalGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, false, (short)11);		
	}

}
