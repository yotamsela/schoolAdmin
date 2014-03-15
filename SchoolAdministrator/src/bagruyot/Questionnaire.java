package bagruyot;
import java.util.ArrayList;
import java.util.List;

public class Questionnaire{
	
	private int number;
	private List<Exam> exams;
	private boolean passed;
	private QuestionnaireGroup group;
	private Exam bestExam;
	private SchoolClass schoolClass;

	public Questionnaire(int number){
		this.number = number;
		exams = new ArrayList<Exam>();
		group = null;
		bestExam = null;
		passed = false;
		schoolClass = null;
	}

	public void addExam(Exam exam, SchoolClass schoolClass) {
		exams.add(exam);
		if(this.schoolClass == null){
			this.schoolClass = schoolClass;
		}
		if(bestExam == null || bestExam.getFinalGrade() < exam.getFinalGrade()){
			bestExam = exam;			
			if(bestExam.getFinalGrade() > Student.PASSING_BAR){
				passed = true;
			}
			if(group != null){
				group.avgChanged();
			}
		}
	}
	
	public SchoolClass getSchoolClass() {
		return schoolClass;
	}

	public void setGroup(QuestionnaireGroup group) {
		this.group = group;
	}

	public boolean isPassed() {
		return passed;
	}

	public Questionnaire copy() {
		return new Questionnaire(number);
	}

	public List<Exam> getExams() {
		return exams;
	}

	public QuestionnaireGroup getGroup() {
		return group;
	}
	
	public Exam getBestExam() {
		return bestExam;
	}

	public int getNumber() {
		return number;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + number;
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		Questionnaire other = (Questionnaire) obj;
		if (number != other.number)
			return false;
		return true;
	}

	@Override
	public String toString() {
		return "Questionnaire [number=" + number + ", exams=" + exams + ", finalGrade=" + "]";
	}

}
