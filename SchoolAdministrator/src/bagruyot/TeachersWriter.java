package bagruyot;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TeachersWriter{

	public static final String FIRST_PAGE_NAME = "ממוצעים";

	public static void write(List<QuestionnaireGroup> groups, String name) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		printStatistics(groups, workbook);
		for (QuestionnaireGroup group : groups) {
			printGroup(group, workbook);
		}
		Writer.write(workbook, name + " - ממוצעי מורים.xlsx");
	}

	private static void printStatistics(List<QuestionnaireGroup> groups, XSSFWorkbook workbook){
		XSSFSheet sheet = workbook.createSheet(FIRST_PAGE_NAME);
		sheet.setRightToLeft(true);
		int rownum = 0;
		XSSFRow row = sheet.createRow(rownum++);
		XSSFCellStyle style;
		Writer.createCell(workbook, row, 0, "מספר קבוצה", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 1, "שנה", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 2, "מקצוע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 3, "מספר שאלון", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 4, "שם המורה", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 5, "מספר תלמידים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 6, "מספר עוברים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 7, "מספר נכשלים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 8, "ציון מגן ממוצע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 9, "ציון בגרות ממוצע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		Writer.createCell(workbook, row, 10, "ציון סופי ממוצע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)11, true, false);
		for (QuestionnaireGroup group : groups) {
			row = sheet.createRow(rownum++);
			Writer.createCell(workbook, row, 0, group.getGroup().getNumber(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, (short)11, IndexedColors.BLUE.getIndex(), IndexedColors.WHITE.getIndex(), 
					true, false, false, true);
			Writer.createCell(workbook, row, 1, group.getGroup().getYear(), XSSFCellStyle.BORDER_THIN,
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 2, Model.subjectsMap.get(Model.questionnaireMap.get(group.getQuestionnaire())).getName(),
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, 
					false, false);
			Writer.createCell(workbook, row, 3, group.getQuestionnaire(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 4, group.getGroup().getTeacher().getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
			Writer.createCell(workbook, row, 5, group.getStudents().size(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 6, group.getNumOfPassed(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 7, group.getStudents().size() - group.getNumOfPassed(), XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 8, group.getAvgMagenGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, 2);
			Writer.createCell(workbook, row, 9, group.getAvgExamGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, 2);
			Writer.createCell(workbook, row, 10, group.getAvgFinalGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, false, (short)11, 2);
		}
		for (int i = 0; i < 11; i++) {
			style = row.getCell(i).getCellStyle();
			Writer.changeCellStyle(workbook, row.getCell(i), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), 
					style.getBorderRight(), 0);
		}
		sheet.autoSizeColumn(0);
		sheet.setColumnWidth(1, 2200);
		sheet.setColumnWidth(2, 4400);
		sheet.setColumnWidth(3, 2500);
		sheet.setColumnWidth(4, 4500);
		sheet.setColumnWidth(5, 2500);
		sheet.setColumnWidth(6, 2500);
		sheet.setColumnWidth(7, 2500);
		sheet.setColumnWidth(8, 2500);
		sheet.setColumnWidth(9, 2500);
		sheet.setColumnWidth(10, 2500);
		sheet.setAutoFilter(CellRangeAddress.valueOf("A1:K"+rownum));
		sheet.createFreezePane(0, 1, 0, 1);
	}

	private static void printGroup(QuestionnaireGroup group, XSSFWorkbook workbook) {
		XSSFSheet sheet = workbook.createSheet(group.getGroup().getNumber() + " - " + group.getGroup().getYear() + " - "
				+ group.getQuestionnaire());
		CreationHelper helper = workbook.getCreationHelper();
		createHyperLink(workbook, sheet, helper, group);
		sheet.setRightToLeft(true);
		printGroupHeader(group, workbook, sheet, helper);
		printGroupSummary(group, workbook, sheet);
		int size = printGroupStudentsHeader(group, workbook, sheet);
		printGroupStudents(group, workbook, sheet, size);
		//TODO fix this
		//		for (int i = 0; i <= size; i++) {
		//			sheet.autoSizeColumn(i);
		//		}
	}

	private static void createHyperLink(XSSFWorkbook workbook, XSSFSheet sheet, CreationHelper helper, QuestionnaireGroup group) {
		XSSFSheet firstSheet = workbook.getSheetAt(0);
		Hyperlink link;
		for (Row row : firstSheet) {
			if(row.getRowNum() > 0 && row.getCell(0).getNumericCellValue() == group.getGroup().getNumber() && 
					row.getCell(1).getNumericCellValue() == group.getGroup().getYear() && 
					row.getCell(3).getNumericCellValue() == group.getQuestionnaire()){
				link = helper.createHyperlink(Hyperlink.LINK_DOCUMENT);
				link.setAddress("'" + group.getGroup().getNumber() + " - " + group.getGroup().getYear() + " - " + group.getQuestionnaire() + 
						"'!A1");
				row.getCell(0).setHyperlink(link);
			}
		}
	}

	private static void printGroupHeader(QuestionnaireGroup group, XSSFWorkbook workbook, XSSFSheet sheet, CreationHelper helper){
		List<SchoolClass> classes = getClasses(group);
		int rownum = 1, numOfClasses = classes.size();
		Hyperlink link;
		XSSFRow row = sheet.createRow(rownum++);
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 1));
		Writer.createCell(workbook, row, 0, "מספר הקבוצה", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		Writer.createCell(workbook, row, 2, "שנה", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 3, 4));
		Writer.createCell(workbook, row, 3, "שם המורה", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);
		Writer.createCell(workbook, row, 4, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 5, 6));
		Writer.createCell(workbook, row, 5, "מספר שאלון", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);
		Writer.createCell(workbook, row, 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 7, 8));
		Writer.createCell(workbook, row, 7, "מקצוע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);
		Writer.createCell(workbook, row, 8, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		sheet.addMergedRegion(new CellRangeAddress(1, 1, 9, 8 + numOfClasses));
		if(numOfClasses == 1){
			Writer.createCell(workbook, row, 9, "כיתה", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		}
		else{
			Writer.createCell(workbook, row, 9, "כיתות אם", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);
			for(int i = 1; i < numOfClasses - 1; i++){
				Writer.createCell(workbook, row, 9 + i, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
						XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)16, false, false);
			}
			Writer.createCell(workbook, row, 8 + numOfClasses, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		}
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));
		Writer.createCell(workbook, row, 9 + numOfClasses, "בחזרה לדף הראשי", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, (short)12, IndexedColors.BLUE.getIndex(), IndexedColors.WHITE.getIndex(), true,
				false, false, true);
		Writer.createCell(workbook, row, 10 + numOfClasses, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)12, false, false);
		link = helper.createHyperlink(Hyperlink.LINK_DOCUMENT);
		link.setAddress("'" + FIRST_PAGE_NAME + "'!A1");
		row.getCell(9 + numOfClasses).setHyperlink(link);
		row = sheet.createRow(rownum++);
		Writer.createCell(workbook, row, 0, group.getGroup().getNumber(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, false, (short)16);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		Writer.createCell(workbook, row, 2, group.getGroup().getYear(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)16);
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 4));
		Writer.createCell(workbook, row, 3, group.getGroup().getTeacher().getName(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, false, (short)16, false, false);
		Writer.createCell(workbook, row, 4, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 6));
		Writer.createCell(workbook, row, 5, group.getQuestionnaire(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, false, (short)16);
		Writer.createCell(workbook, row, 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		sheet.addMergedRegion(new CellRangeAddress(2, 2, 7, 8));
		Writer.createCell(workbook, row, 7, Model.subjectsMap.get(Model.questionnaireMap.get(group.getQuestionnaire())).getName(), 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, false, (short)16, 
				false, false);
		Writer.createCell(workbook, row, 8, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
		if(numOfClasses == 1){
			Writer.createCell(workbook, row, 9, classes.get(0).getName(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		}
		else{
			Writer.createCell(workbook, row, 9, classes.get(0).getName(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
			for(int i = 1; i < numOfClasses - 1; i++){
				Writer.createCell(workbook, row, 9 + i, classes.get(i).getName(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)16, false, false);
			}
			Writer.createCell(workbook, row, 8 + numOfClasses, classes.get(numOfClasses - 1).getName(), XSSFCellStyle.BORDER_MEDIUM,
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		}
		Writer.createCell(workbook, row, 9 + numOfClasses, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, true, (short)16, false, false);
		Writer.createCell(workbook, row, 10 + numOfClasses, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, true, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(1, 2, 9 + numOfClasses, 10 + numOfClasses));
	}

	private static void printGroupSummary(QuestionnaireGroup group, XSSFWorkbook workbook, XSSFSheet sheet){
		int rownum = 4, numOfStudents = group.getStudents().size(), numOfPassed = group.getNumOfPassed();
		XSSFRow row = sheet.createRow(rownum++);
		sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 1));
		Writer.createCell(workbook, row, 0, "מספר התלמידים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, true, (short)12, false, false);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(4, 4, 2, 3));
		Writer.createCell(workbook, row, 2, "מספר עוברים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)12, false, false);
		Writer.createCell(workbook, row, 3, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)12, false, false);
		Writer.createCell(workbook, row, 4, "אחוז עוברים", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(4, 4, 5, 6));
		Writer.createCell(workbook, row, 5, "ציון מגן ממוצע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)12, false, false);
		Writer.createCell(workbook, row, 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(4, 4, 7, 8));
		Writer.createCell(workbook, row, 7, "ציון בגרות ממוצע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)12, false, false);
		Writer.createCell(workbook, row, 8, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, true, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(4, 4, 9, 10));
		Writer.createCell(workbook, row, 9, "ציון סופי ממוצע", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, true, (short)12, false, false);
		Writer.createCell(workbook, row, 10, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)12, false, false);
		row = sheet.createRow(rownum++);
		sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));
		Writer.createCell(workbook, row, 0, numOfStudents, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, false, (short)12);
		Writer.createCell(workbook, row, 1, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, false, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(5, 5, 2, 3));
		Writer.createCell(workbook, row, 2, numOfPassed, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)12);
		Writer.createCell(workbook, row, 3, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, false, (short)12, false, false);
		Writer.createCell(workbook, row, 4, 100.0 * numOfPassed / numOfStudents, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, false, (short)12, 2);
		sheet.addMergedRegion(new CellRangeAddress(5, 5, 5, 6));
		Writer.createCell(workbook, row, 5, group.getAvgMagenGrade(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, false, (short)12, 2);
		Writer.createCell(workbook, row, 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, false, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(5, 5, 7, 8));
		Writer.createCell(workbook, row, 7, group.getAvgExamGrade(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, false, (short)12, 2);
		Writer.createCell(workbook, row, 8, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_THIN, false, (short)12, false, false);
		sheet.addMergedRegion(new CellRangeAddress(5, 5, 9, 10));
		Writer.createCell(workbook, row, 9, group.getAvgFinalGrade(), XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_NONE, false, (short)12, 2);
		Writer.createCell(workbook, row, 10, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, false, (short)12, false, false);
	}

	private static int printGroupStudentsHeader(QuestionnaireGroup group, XSSFWorkbook workbook, XSSFSheet sheet) {
		XSSFRow upperRow = sheet.createRow(7), lowerRow = sheet.createRow(8);
		sheet.addMergedRegion(new CellRangeAddress(7, 8, 0, 0));
		Writer.createCell(workbook, upperRow, 0, "כיתת אם", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, lowerRow, 0, "", XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		sheet.addMergedRegion(new CellRangeAddress(7, 8, 1, 1));
		Writer.createCell(workbook, upperRow, 1, "שם התלמיד", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, lowerRow, 1, "", XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		sheet.addMergedRegion(new CellRangeAddress(7, 8, 2, 2));
		Writer.createCell(workbook, upperRow, 2, "מספר זהות", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		Writer.createCell(workbook, lowerRow, 2, "", XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, true, (short)11, false, false);
		sheet.addMergedRegion(new CellRangeAddress(7, 8, 3, 3));
		Writer.createCell(workbook, upperRow, 3, "עבר / נכשל", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, true, false);
		Writer.createCell(workbook, lowerRow, 3, "", XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
		sheet.addMergedRegion(new CellRangeAddress(7, 7, 4, 7));
		Writer.createCell(workbook, upperRow, 4, "בחינה קובעת", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, true, (short)11, true, false);
		Writer.createCell(workbook, upperRow, 5, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
		Writer.createCell(workbook, upperRow, 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
		Writer.createCell(workbook, upperRow, 7, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
		Writer.createCell(workbook, lowerRow, 4, "מועד הבחינה", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, true, (short)11, true, false);
		Writer.createCell(workbook, lowerRow, 5, "ציון מגן", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
		Writer.createCell(workbook, lowerRow, 6, "ציון בחינה", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
		Writer.createCell(workbook, lowerRow, 7, "ציון סופי", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
				XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
		int maxNumOfExams = 0;
		List<Student> students = group.getStudents();
		for (Student student : students) {
			maxNumOfExams = Math.max(maxNumOfExams, student.getQuestionnaire(group).getExams().size());
		}
		for (int i = 1; i < maxNumOfExams; i++) {
			sheet.addMergedRegion(new CellRangeAddress(7, 7, 4 * i + 4, 4 * i + 7));
			Writer.createCell(workbook, upperRow, 4 * i + 4, "בחינה נוספת", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, true, (short)11, true, false);
			Writer.createCell(workbook, upperRow, 4 * i + 5, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
			Writer.createCell(workbook, upperRow, 4 * i + 6, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
			Writer.createCell(workbook, upperRow, 4 * i + 7, "", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
			Writer.createCell(workbook, lowerRow, 4 * i + 4, "מועד הבחינה", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_NONE, true, (short)11, true, false);
			Writer.createCell(workbook, lowerRow, 4 * i + 5, "ציון מגן", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
			Writer.createCell(workbook, lowerRow, 4 * i + 6, "ציון בחינה", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_NONE, true, (short)11, false, false);
			Writer.createCell(workbook, lowerRow, 4 * i + 7, "ציון סופי", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, 
					XSSFCellStyle.BORDER_NONE, XSSFCellStyle.BORDER_MEDIUM, true, (short)11, false, false);
		}
		return 4 * maxNumOfExams + 3;
	}

	private static void printGroupStudents(QuestionnaireGroup group, XSSFWorkbook workbook, XSSFSheet sheet, int size){
		List<Student> students = group.getStudents();
		int rowNum = 9, maxNumOfExams = 0, index;
		for (Student student : students) {
			maxNumOfExams = Math.max(maxNumOfExams, student.getQuestionnaire(group).getExams().size());
		}
		XSSFRow row = null;
		XSSFCellStyle style;
		Questionnaire questionnaire;
		List<Exam> exams;
		Exam bestExam;
		for (Student student : students) {
			row = sheet.createRow(rowNum++);
			questionnaire = student.getQuestionnaire(group);
			bestExam = questionnaire.getBestExam();
			Writer.createCell(workbook, row, 0, questionnaire.getSchoolClass().getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
			Writer.createCell(workbook, row, 1, student.getName(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
			Writer.createCell(workbook, row, 2, student.getIdNumber(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			if(questionnaire.isPassed()){
				Writer.createCell(workbook, row, 3, "עבר", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_MEDIUM, (short)11, IndexedColors.BLACK.getIndex(), IndexedColors.GREEN.getIndex(), false, false, false, 
						false);
			}
			else{
				Writer.createCell(workbook, row, 3, "נכשל", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_MEDIUM, (short)11, IndexedColors.BLACK.getIndex(), IndexedColors.RED.getIndex(), false, false, false, false);
			}
			Writer.createCell(workbook, row, 4, bestExam.getYear(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 5, bestExam.getMagenGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 6, bestExam.getExamGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
			Writer.createCell(workbook, row, 7, bestExam.getFinalGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
					XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, false, (short)11);
			exams = questionnaire.getExams();
			index = 8;
			for (Exam exam : exams) {
				if(!exam.equals(bestExam)){
					Writer.createCell(workbook, row, index++, exam.getYear(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, false, (short)11);
					Writer.createCell(workbook, row, index++, exam.getMagenGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
					Writer.createCell(workbook, row, index++, exam.getExamGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11);
					Writer.createCell(workbook, row, index++, exam.getFinalGrade(), XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
							XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, false, (short)11);		
				}
			}
			while(index <= size){
				Writer.createCell(workbook, row, index++, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
				Writer.createCell(workbook, row, index++, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
				Writer.createCell(workbook, row, index++, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, false, (short)11, false, false);
				Writer.createCell(workbook, row, index++, "", XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_THIN, 
						XSSFCellStyle.BORDER_THIN, XSSFCellStyle.BORDER_MEDIUM, false, (short)11, false, false);
			}
		}
		for (int i = 0; i <= size; i++) {
			style = row.getCell(i).getCellStyle();
			Writer.changeCellStyle(workbook, row.getCell(i), style.getBorderTop(), XSSFCellStyle.BORDER_MEDIUM, style.getBorderLeft(), style.getBorderRight(),
				0);
		}
		sheet.setColumnWidth(1, 3500);
		sheet.setColumnWidth(4, 2800);
	}

	private static List<SchoolClass> getClasses(QuestionnaireGroup group) {
		Set<SchoolClass> tempClasses = new HashSet<SchoolClass>();
		List<SchoolClass> classes = new ArrayList<SchoolClass>();
		List<Student> students = group.getStudents();
		for (Student student : students) {
			tempClasses.add(student.getQuestionnaire(group).getSchoolClass());
		}
		for (Iterator<SchoolClass> iterator = tempClasses.iterator(); iterator.hasNext();) {
			classes.add(iterator.next());
			Collections.sort(classes);
		}
		return classes;
	}

}
