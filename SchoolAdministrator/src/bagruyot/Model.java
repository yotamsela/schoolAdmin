package bagruyot;
import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.swing.JOptionPane;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Model implements Runnable {

	//TODO change to non static
	public static Map<Subject, List<Map<Questionnaire, Double>>> percentagesMap = new HashMap<Subject, List<Map<Questionnaire, Double>>>();
	public static Map<Questionnaire, Subject> subjectsMap = new HashMap<Questionnaire, Subject>();
	public static Map<Integer, Questionnaire> questionnaireMap = new HashMap<Integer, Questionnaire>();
	public static Map<String, Teacher> teachersMap = new HashMap<String, Teacher>();
	public static Map<Integer, Map<String, SchoolClass>> classesMap = new HashMap<Integer, Map<String, SchoolClass>>();// year => class name => group
	public static Map<Integer, Map<Integer, Group>> groupsMap = new HashMap<Integer, Map<Integer, Group>>();// year => group number => group
	public static Map<String, Student> studentsMap = new HashMap<String, Student>();
	public static Map<Subject, Double> bonusesMap = new HashMap<Subject, Double>();
	public static Set<Integer> unrecognizedQuestionnaires = new HashSet<Integer>();
	public static Set<Integer> unrecognizedGroups = new HashSet<Integer>();
	
	private boolean writeGrade10;
	private boolean writeGrade11;
	private boolean writeGrade12;
	private boolean writeTeachers;
	private boolean writeStudents;
	private boolean writeErrors;
	private String filePath;
	
	public Model(boolean writeGrade10, boolean writeGrade11, boolean writeGrade12, boolean writeTeachers, 
			boolean writeStudents, boolean writeErrors, String filePath){
		this.writeGrade10 = writeGrade10;
		this.writeGrade11 = writeGrade11;
		this.writeGrade12 = writeGrade12;
		this.writeTeachers = writeTeachers;
		this.writeStudents = writeStudents;
		this.writeErrors = writeErrors;
		this.filePath = filePath;
	}
	
	@Override
	public void run() {
		fullCalculation(writeGrade10, writeGrade11, writeGrade12, writeTeachers, writeStudents, writeErrors, filePath);
	}

	public static void fullCalculation(boolean writeGrade10, boolean writeGrade11, boolean writeGrade12, boolean writeTeachers, 
			boolean writeStudents, boolean writeErrors, String filePath){
		SchoolClass.fillMaps();
		File file = new File(filePath);
		String fileName = filePath.substring(0, filePath.indexOf('.'));
		XSSFSheet sheet = StudentsReader.fileAnalysis(file);
		if(sheet == null){
			//TODO add error message
			return;
		}
		List<QuestionnaireGroup> qGroups = new ArrayList<QuestionnaireGroup>();
		Set<Integer> groupNumbers;
		Set<Integer> years = groupsMap.keySet();
		for (Integer year : years) {
			groupNumbers = groupsMap.get(year).keySet();
			for (Integer groupNumber : groupNumbers) {
				qGroups.addAll(groupsMap.get(year).get(groupNumber).getQuestionnaires());
			}
		}
		if(writeTeachers){
			TeachersWriter.write(qGroups, fileName);			
		}
		if(writeStudents && (writeGrade10 || writeGrade11 || writeGrade12)){
			List<Subject> subjects = new ArrayList<Subject>();
			Subject tempSubject, subject;
			List<Map<Questionnaire, Double>> tempList;
			Set<Double> unitsNumbers;
			Set<Subject> tempSubjects = percentagesMap.keySet();
			for (Iterator<Subject> subjectsIterator = tempSubjects.iterator(); subjectsIterator.hasNext();) {
				tempSubject = subjectsIterator.next();
				tempList = percentagesMap.get(tempSubject);
				unitsNumbers = new HashSet<Double>();
				for (Map<Questionnaire, Double> map : tempList) {
					unitsNumbers.add(map.get(null));
				}
				for (Iterator<Double> unitsIterator = unitsNumbers.iterator(); unitsIterator.hasNext();) {
					subject = tempSubject.copy();
					subject.setNumOfUnits((int)Math.round(unitsIterator.next()));
					subjects.add(subject);
				}
			}
			Collections.sort(subjects);
			Set<String> tempS = studentsMap.keySet();
			List<Student> students = new ArrayList<Student>();
			Student student;
			int grade;
			for (Iterator<String> iterator = tempS.iterator(); iterator.hasNext();) {
				student = studentsMap.get(iterator.next());
				grade = student.getSchoolGrade();
				if((writeGrade10 && grade == 10) || (writeGrade11 && grade == 11) || (writeGrade12 && grade == 12)){
					students.add(student);
				}
			}
			Collections.sort(students);
			boolean flag;
			Subject sub;
			for (int i = 0; i < subjects.size(); i++) {
				sub = subjects.get(i);
				flag = false;
				for (Student stu: students) {
					if(stu.getRelevantSubjects().contains(sub)){
						flag = true;
					}
				}
				if(!flag){
					subjects.remove(sub);
				}
			}
			StudentsWriter.write(subjects, students, fileName);
		}
		if(writeErrors){
			if(unrecognizedQuestionnaires.size() > 0){				
				unrecognizedQuestionnaires.add(0);
			}
			if(unrecognizedGroups.size() > 0){				
				unrecognizedGroups.add(0);
			}			
			if(unrecognizedQuestionnaires.size() > 0 || unrecognizedGroups.size() > 0){				
				ErrorsWriter.write(sheet, unrecognizedQuestionnaires, unrecognizedGroups,  fileName);
			}			
		}
		JOptionPane.showMessageDialog(null, "החישוב הושלם");
	}

}
