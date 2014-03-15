package bagruyot;
import java.util.ArrayList;
import java.util.List;

public class Group {

	private int number;
	private int year;
	private Teacher teacher;
	private List<QuestionnaireGroup> questionnaires;

	public Group(int number, Teacher teacher, int year) {
		this.number = number;
		this.teacher = teacher;
		this.year = year;
		questionnaires = new ArrayList<QuestionnaireGroup>();
	}

	public void addQuestionnaireGroup(int qNumber){
		for (QuestionnaireGroup q : questionnaires) {
			if(q.getQuestionnaire() == qNumber){
				return;
			}
		}
		questionnaires.add(new QuestionnaireGroup(qNumber, this));
	}
	
	public int getNumber() {
		return number;
	}

	public Teacher getTeacher() {
		return teacher;
	}

	public QuestionnaireGroup getQuestionnaireGroup(int qNumber) {
		for (QuestionnaireGroup q : questionnaires) {
			if(q.getQuestionnaire() == qNumber){
				return q;
			}
		}
		return null;
	}


	public List<QuestionnaireGroup> getQuestionnaires() {
		return questionnaires;
	}

	public int getYear() {
		return year;
	}

}
