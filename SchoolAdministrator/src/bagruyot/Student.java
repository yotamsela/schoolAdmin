package bagruyot;
import java.util.ArrayList;
import java.util.List;

public class Student implements Comparable<Student>{

	public static final int MIN_NUM_OF_UNITS = 20;
	public static final int PASSING_BAR = 55;

	private int idNumber;
	private double finalGrade;	
	private double weightedGrade;
	private int numOfUnits;
	private String name;
	private List<Subject> subjects;
	private SchoolClass schoolClass;
	private boolean isEntitled;
	private boolean avgCalculated;
	private List<Subject> relevantSubjects;
	
	public Student(String name, int idNumber) {
		this.name = name;
		this.idNumber = idNumber;
		subjects = new ArrayList<Subject>();
		relevantSubjects = new ArrayList<Subject>();
		avgCalculated = false;
	}

	public void addExam(int examNumber, int magenGrade, int examGrade, int finalGrade, int date, QuestionnaireGroup group, SchoolClass schoolClass) {
		avgCalculated = false;
		Exam exam = new Exam(magenGrade, examGrade, finalGrade, date);
		if(this.schoolClass == null || (schoolClass != null && this.schoolClass.getName().compareTo(schoolClass.getName()) < 0)){
			this.schoolClass = schoolClass;
		}
		Questionnaire newQuestionnaire = Model.questionnaireMap.get(examNumber).copy();
		Subject newSubject = Model.subjectsMap.get(newQuestionnaire).copy();		
		for (Subject sub : subjects) {
			if(sub.equals(newSubject)){
				for (Questionnaire ques : sub.getQuestionnaires()) {
					if(ques.equals(newQuestionnaire)){
						if(group != null && newQuestionnaire.getGroup() == null){
							ques.setGroup(group);
						}
						ques.addExam(exam, schoolClass);
						return;
					}
				}
				if(group != null && newQuestionnaire.getGroup() == null){
					newQuestionnaire.setGroup(group);			
				}
				newQuestionnaire.addExam(exam, schoolClass);
				sub.addQuestionnaire(newQuestionnaire, schoolClass);
				return;
			}
		}
		if(group != null && newQuestionnaire.getGroup() == null){
			newQuestionnaire.setGroup(group);			
		}		
		newQuestionnaire.addExam(exam, schoolClass);
		newSubject.addQuestionnaire(newQuestionnaire, schoolClass);
		subjects.add(newSubject);
	}
	
	public Questionnaire getQuestionnaire(QuestionnaireGroup group){
		Questionnaire questionnaire = Model.questionnaireMap.get(group.getQuestionnaire());
		for (Subject subject : subjects) {
			for (Questionnaire q : subject.getQuestionnaires()) {
				if(questionnaire.equals(q)){
					return q;
				}
			}
		}
		return null;
	}
	
	private void calcAvg(){
		finalGrade = 0;
		weightedGrade = 0;
		numOfUnits = 0;
		isEntitled = true;
		relevantSubjects.clear();
		for (Subject subject : subjects) {
			if(subject.isMandatory() || subject.isCompleated()){
				relevantSubjects.add(subject);
			}
			if(subject.isMandatory() && !subject.hasPassed()){
				isEntitled = false;
			}
		}
		for (Subject subject : relevantSubjects) {
			numOfUnits += subject.getNumOfUnits();
			finalGrade += subject.getFinalGrade() * subject.getNumOfUnits();
			weightedGrade += (subject.getFinalGrade() + Model.bonusesMap.get(subject)) * subject.getNumOfUnits();
		}
		for (Subject subject : relevantSubjects) {
			if(!subject.isMandatory() && !subject.hasPassed() && numOfUnits - subject.getNumOfUnits() >= MIN_NUM_OF_UNITS){
				relevantSubjects.remove(subject);
				finalGrade += subject.getFinalGrade() * subject.getNumOfUnits();
				weightedGrade += (subject.getFinalGrade() + Model.bonusesMap.get(subject)) * subject.getNumOfUnits();
			}
		}		
		if(!isEntitled){
			isEntitled = calcSpecialCases();
		}
		isEntitled = isEntitled && numOfUnits >= MIN_NUM_OF_UNITS;
		finalGrade /= numOfUnits;
		weightedGrade /= numOfUnits;
		avgCalculated = true;
	}
	
	//TODO make sure it is the right rule. magic numbers?
	public boolean calcSpecialCases(){
		List<Subject> failedSubjects = new ArrayList<Subject>();
		for (Subject subject : relevantSubjects) {
			if(!subject.hasPassed()){
				failedSubjects.add(subject);
			}
		}
		Subject fail = failedSubjects.get(0), best1 = null, best2 = null;
		if(failedSubjects.size() > 1 || !fail.isCompleated() || fail.getName().equals("לשון")){
			return false;
		}
		if(fail.getFinalGrade() > 45){
			return true;
		}
		for (Subject subject : relevantSubjects) {
			if(subject.getNumOfUnits() >= 3){
				if(best2 == null || subject.getFinalGrade() > best2.getFinalGrade()){
					best2 = subject;
				}
				if(best1 == null || subject.getFinalGrade() > best1.getFinalGrade()){
					best2 = best1;
					best1 = subject;
				}
			}
		}
		if(best1 != null && best2 != null && best1.getFinalGrade() + best2.getFinalGrade() >= 150){
			return true;
		}
		return false;
	}

	public int getNumOfUnits() {
		if(!avgCalculated){
			calcAvg();
		}
		return numOfUnits;
	}

	public String getName() {
		return name;
	}

	public int getIdNumber() {
		return idNumber;
	}

	public double getFinalGrade() {
		if(!avgCalculated){
			calcAvg();
		}
		return finalGrade;
	}

	public SchoolClass getSchoolClass() {
		return schoolClass;
	}

	public int getSchoolGrade() {
		String className = schoolClass.getName();
		return SchoolClass.classesMap.get(className.substring(0, className.length() - 2));
	}

	public boolean isEntitled() {
		if(!avgCalculated){
			calcAvg();
		}
		return isEntitled;
	}

	public double getWeightedGrade() {
		if(!avgCalculated){
			calcAvg();
		}
		return weightedGrade;
	}

	public List<Subject> getRelevantSubjects() {
		if(!avgCalculated){
			calcAvg();
		}
		return relevantSubjects;
	}

//	@Override
//	public String toString() {
//		return "Student [name=" + name + ", subjects=" + subjects + "]";
//	}

	@Override
	public int compareTo(Student other) {
		return name.compareTo(other.name);
	}

}
