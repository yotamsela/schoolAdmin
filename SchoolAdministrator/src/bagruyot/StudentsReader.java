package bagruyot;
import java.io.File;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class StudentsReader {
	public static final String STUDENT_NAME = "שם תלמיד";
	public static final String QUESTIONNAIRE_NAME = "סמל שאלון";
	public static final String FINAL_GRADE_NAME = "ציון סופי";
	public static final String MAGEN_GRADE_NAME = "ציון שנתי/מגן";
	public static final String EXAM_GRADE_NAME = "ציון בחינה";
	public static final String DATE_NAME = "מועד";
	public static final String GROUP_NAME = "קבוצת לימוד";
	public static final String CLASS_NAME = "כיתה";
	public static final String TEACHER_NAME = "מורה";
	public static final String ID_NAME = "ת.ז. תלמיד";
	public static Map<String, Integer> columnNumbers = new HashMap<String, Integer>();

	public static XSSFSheet fileAnalysis(File file){
		XSSFWorkbook oldWorkbook = ExcellReaderAndWriter.openXlsxFile(file);
		if(oldWorkbook == null){
			return null;
		}
		Model.classesMap.clear();
		Model.teachersMap.clear();
		Model.groupsMap.clear();
		Model.studentsMap.clear();
		StudentsReader.columnNumbers.clear();
		XSSFSheet firstSheet = oldWorkbook.getSheetAt(0);
		if(!fillColumnNumbers(firstSheet)){
			return null;
		}
		checkSubjectMaps(firstSheet);
		fillGroupsMap(firstSheet);
		fillStudentsMap(firstSheet);
//		List<Student> students = new ArrayList<Student>();
//		Set<String> studentNames = SchoolProgram.studentsMap.keySet();
//		for (String string : studentNames) {
//			students.add(SchoolProgram.studentsMap.get(string));
//		}
//		Collections.sort(students);
		return firstSheet;
	}

	private static boolean fillColumnNumbers(XSSFSheet firstSheet) {
		XSSFRow firstRow = firstSheet.getRow(0);
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, STUDENT_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, QUESTIONNAIRE_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, MAGEN_GRADE_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, EXAM_GRADE_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, FINAL_GRADE_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, DATE_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, GROUP_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, CLASS_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, TEACHER_NAME)){
			return false;
		}
		if(!addIndexToMap(firstRow, StudentsReader.columnNumbers, ID_NAME)){
			return false;
		}
		return true;
	}

	private static boolean addIndexToMap(XSSFRow firstRow, Map<String, Integer> columnNumbers, String name){
		Iterator<Cell> cellIterator = firstRow.iterator();
		Cell cell;
		int i = 0, nameIndex = -1;
		while(cellIterator.hasNext()){
			cell = cellIterator.next();
			if(cell.toString().equals(name) || cell.toString().substring(1).equals(name)){
				nameIndex = i;
				break;
			}
			i++;
		}
		if(nameIndex == -1){
			JOptionPane.showMessageDialog(null, "תקלה: התא - " + name + "לא נמצא בשורה העליונה של הגיליון הראשון", "שגיאת פורמט בדף", JOptionPane.ERROR_MESSAGE);
			return false;
		}
		columnNumbers.put(name, nameIndex);
		return true;
	}

	private static void checkSubjectMaps(XSSFSheet firstSheet) {
		int questionnaireIndex = StudentsReader.columnNumbers.get(QUESTIONNAIRE_NAME);
		for (Row row : firstSheet) {
			if(row.getRowNum() > 0){
				if(row.getCell(questionnaireIndex) != null && row.getCell(questionnaireIndex).toString() != ""){
					if(Model.questionnaireMap.get((int) row.getCell(questionnaireIndex).getNumericCellValue()) == null){
						//TODO should print an error
					}
				}
			}
		}
	}

	private static void fillGroupsMap(XSSFSheet firstSheet) {
		int groupIndex = StudentsReader.columnNumbers.get(GROUP_NAME), teacherIndex = StudentsReader.columnNumbers.get(TEACHER_NAME), 
				questionnaireIndex = StudentsReader.columnNumbers.get(QUESTIONNAIRE_NAME), dateIndex = StudentsReader.columnNumbers.get(DATE_NAME),
				classIndex = StudentsReader.columnNumbers.get(CLASS_NAME);
		int groupNumber, qNumber, year;
		String teacherName, className;
		Map<String, SchoolClass> yearClasses;
		Teacher teacher;
		Group group;
		Map<Integer, Group> yearGroups;
		SchoolClass schoolClass;
		for (Row row : firstSheet) {
			if(row.getRowNum() > 0){
				qNumber = (int) row.getCell(questionnaireIndex).getNumericCellValue();
				//TODO change to error message
				if(Model.questionnaireMap.get(qNumber) == null){
					Model.unrecognizedQuestionnaires.add(row.getRowNum());
					continue;
				}
				if(row.getCell(teacherIndex) != null && row.getCell(teacherIndex).toString().length() > 0){
					teacherName = row.getCell(teacherIndex).toString();
					teacher = Model.teachersMap.get(teacherName);
					if(teacher == null){
						teacher = new Teacher(teacherName);
						Model.teachersMap.put(teacherName, teacher);
					}
					groupNumber = (int) row.getCell(groupIndex).getNumericCellValue();
					year = (int) row.getCell(dateIndex).getNumericCellValue() / 100;
					yearGroups = Model.groupsMap.get(year);
					if(yearGroups == null){
						yearGroups = new HashMap<Integer, Group>();
						Model.groupsMap.put(year, yearGroups);
					}
					group = yearGroups.get(groupNumber);
					if(group == null){
						group = new Group(groupNumber, teacher, year);
						yearGroups.put(groupNumber, group);
					}
					group.addQuestionnaireGroup(qNumber);
					teacher.addGroup(group, year);
				}
				if(row.getCell(classIndex) != null && row.getCell(classIndex).toString().length() > 0){
					year = (int) row.getCell(dateIndex).getNumericCellValue() / 100;
					className = row.getCell(classIndex).toString();
					yearClasses = Model.classesMap.get(year);
					if(yearClasses == null){
						yearClasses = new HashMap<String, SchoolClass>();
						Model.classesMap.put(year, yearClasses);
					}
					schoolClass = yearClasses.get(className);
					if(schoolClass == null){
						schoolClass = new SchoolClass(className, year);
						yearClasses.put(className, schoolClass);
					}
				}					
			}
		}
		Model.groupsMap.remove(0);
	}

	private static void fillStudentsMap(XSSFSheet firstSheet) {
		int nameIndex = StudentsReader.columnNumbers.get(STUDENT_NAME), questionnaireIndex = StudentsReader.columnNumbers.get(QUESTIONNAIRE_NAME), 
				magenGradeIndex = StudentsReader.columnNumbers.get(MAGEN_GRADE_NAME), examGradeIndex = StudentsReader.columnNumbers.get(EXAM_GRADE_NAME), 
				finalGradeIndex = StudentsReader.columnNumbers.get(FINAL_GRADE_NAME), dateIndex = StudentsReader.columnNumbers.get(DATE_NAME),
				groupIndex = StudentsReader.columnNumbers.get(GROUP_NAME), idIndex = StudentsReader.columnNumbers.get(ID_NAME),
				classIndex = StudentsReader.columnNumbers.get(CLASS_NAME);
		String name, className;
		Student student;
		int qNumber, magenGrade, examGrade, finalGrade, year, groupNumber, id, date;
		SchoolClass schoolClass;
		QuestionnaireGroup qGroup = null;
		for (Row row : firstSheet) {
			if(row.getRowNum() > 0){
				qNumber = (int) row.getCell(questionnaireIndex).getNumericCellValue();
				//TODO change to error message
				if(Model.questionnaireMap.get(qNumber) == null){
					Model.unrecognizedQuestionnaires.add(row.getRowNum());
					continue;
				}
				name = row.getCell(nameIndex).toString();
				id = (int) row.getCell(idIndex).getNumericCellValue();
				date = (int) row.getCell(dateIndex).getNumericCellValue();
				year = date / 100;
				student = Model.studentsMap.get(name);
				if(student == null){
					student = new Student(name, id);					
					Model.studentsMap.put(name, student);
				}
				schoolClass = null;
				if(row.getCell(classIndex) != null && row.getCell(classIndex).toString().length() > 0){
					className = row.getCell(classIndex).toString();
					schoolClass = Model.classesMap.get(year).get(className);
					schoolClass.addStudent(student);
				}
				qGroup = null;
				if(row.getCell(groupIndex) != null && row.getCell(groupIndex).toString().length() > 0){
					groupNumber = (int)row.getCell(groupIndex).getNumericCellValue();
					if(Model.groupsMap.get(year).get(groupNumber) != null){
						qGroup = Model.groupsMap.get(year).get(groupNumber).getQuestionnaireGroup(qNumber);
						qGroup.addStudent(student, schoolClass, year);						
					}
					else{
						Model.unrecognizedGroups.add(row.getRowNum());
					}
				}
				magenGrade = (row.getCell(magenGradeIndex) == null || row.getCell(magenGradeIndex).toString().equals(""))? 0: 
					(int) row.getCell(magenGradeIndex).getNumericCellValue();
				examGrade = (row.getCell(examGradeIndex) == null || row.getCell(examGradeIndex).toString().equals(""))? 0: 
					(int) row.getCell(examGradeIndex).getNumericCellValue();
				finalGrade = (row.getCell(finalGradeIndex) == null || row.getCell(finalGradeIndex).toString().equals(""))? 0: 
					(int) row.getCell(finalGradeIndex).getNumericCellValue();
				student.addExam(qNumber, magenGrade, examGrade, finalGrade, date, qGroup, schoolClass);
			}
		}
	}

}
