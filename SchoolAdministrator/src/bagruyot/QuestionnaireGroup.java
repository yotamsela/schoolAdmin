package bagruyot;
import java.util.ArrayList;
import java.util.List;

public class QuestionnaireGroup {

	private Group group;
	private int questionnaire;
	private List<Student> students;
	private double avgMagenGrade;
	private double avgExamGrade;
	private double avgFinalGrade;
	private int numOfPassed;
	private boolean avgCalculated;

	public QuestionnaireGroup(int qNumber, Group group){
		students = new ArrayList<Student>();
		this.group = group;
		questionnaire = qNumber;
		avgCalculated = false;
	}

	public void addStudent(Student newStudent, SchoolClass schoolClass, int year) {
		for (Student student: students) {
			if(newStudent.equals(student)){
				return;
			}
		}
		students.add(newStudent);
		avgCalculated = false;
	}

	private void calcAvg(){
		avgMagenGrade = 0;
		avgExamGrade = 0;
		avgFinalGrade = 0;
		numOfPassed = 0;
		Questionnaire questionnaire;
		for (Student student : students) {
			questionnaire = student.getQuestionnaire(this);
			avgMagenGrade += questionnaire.getBestExam().getMagenGrade();
			avgExamGrade += questionnaire.getBestExam().getExamGrade();
			avgFinalGrade += questionnaire.getBestExam().getFinalGrade();
			if(questionnaire.isPassed()){
				numOfPassed++;
			}
		}
		avgMagenGrade /= students.size();
		avgExamGrade /= students.size();
		avgFinalGrade /= students.size();
	}

	public void avgChanged(){
		avgCalculated = false;		
	}
	
	public double getAvgMagenGrade() {
		if(!avgCalculated){
			calcAvg();
			avgCalculated = true;
		}
		return avgMagenGrade;
	}

	public double getAvgExamGrade() {
		if(!avgCalculated){
			calcAvg();
			avgCalculated = true;
		}
		return avgExamGrade;
	}

	public double getAvgFinalGrade() {
		if(!avgCalculated){
			calcAvg();
			avgCalculated = true;
		}
		return avgFinalGrade;
	}

	public int getQuestionnaire() {
		return questionnaire;
	}

	public Group getGroup() {
		return group;
	}

	public List<Student> getStudents() {
		return students;
	}
	
	public int getNumOfPassed() {
		if(!avgCalculated){
			calcAvg();
			avgCalculated = true;
		}
		return numOfPassed;
	}

}
