package bagruyot;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class Subject implements Comparable<Subject>{

	private String name;
	private int numOfUnits;
	private int finalGrade;
	private double percentage;
	private Map<Questionnaire, Double> relevantQuestionaires;
	private List<Questionnaire> questionnaires;
	private boolean isMandatory;
	private boolean avgCalculated;
	
	public Subject(String name, boolean isMandatory, int numOfUnits){
		this.name = name;
		this.isMandatory = isMandatory;
		this.numOfUnits = numOfUnits;
		finalGrade = 0;
		percentage = 0.0;
		questionnaires = new ArrayList<Questionnaire>();
		relevantQuestionaires = null;
		avgCalculated = false;
	}

	public void addQuestionnaire(Questionnaire questionnaire, SchoolClass schoolClass) {
		//TODO temporary
		if(questionnaire.getExams().size() == 0){
			questionnaire.getBestExam().getExamGrade();
		}
		for (Questionnaire ques : questionnaires) {
			if(ques.equals(questionnaire)){
				List<Exam> exams = questionnaire.getExams();
				for (Exam exam : exams) {
					ques.addExam(exam, schoolClass);
				}
				return;
			}
		}
		questionnaires.add(questionnaire);
		avgCalculated = false;
	}

	private void checkFullSubject() {
		List<Map<Questionnaire, Double>> options = Model.percentagesMap.get(this);
		Iterator<Questionnaire> it;
		double percentage;
		Questionnaire questionaire;
		for (Map<Questionnaire, Double> map: options) {
			percentage = 0.0;
			for (it = map.keySet().iterator(); it.hasNext();) {
				questionaire = it.next();
				percentage += (this.questionnaires.contains(questionaire))? map.get(questionaire): 0;
			}
			if(percentage > this.percentage){
				this.percentage = percentage;
				relevantQuestionaires = map;
				this.numOfUnits = (int) Math.round(map.get(null));
				if(percentage > 99.0){
					finalGrade = calcGrade(map);
				}
			}
			else if((percentage > 99.0) && ((map.get(null) > this.numOfUnits) || ((map.get(null) == this.numOfUnits) && 
					calcGrade(map) > finalGrade))){
				this.percentage = percentage;
				relevantQuestionaires = map;
				this.numOfUnits = (int) Math.round(map.get(null));
				finalGrade = calcGrade(map);
			}
		}
		avgCalculated = true;
	}

	private int calcGrade(Map<Questionnaire, Double> map) {
		double output = 0.0;
		Questionnaire q1;
		for (Iterator<Questionnaire> it = map.keySet().iterator(); it.hasNext();) {
			q1 = it.next();
			for(Questionnaire q2: questionnaires){
				if(q1 != null && q1.equals(q2)){
					output += q2.getBestExam().getFinalGrade() * map.get(q1) / 100.0;
					break;
				}
			}
		}
		return (int) Math.round(output);
	}

	public List<Questionnaire> getQuestionnaires() {
		return questionnaires;
	}

	public int getNumOfUnits() {
		if(!avgCalculated){
			checkFullSubject();
		}
		return numOfUnits;
	}

	public boolean isCompleated() {
		if(!avgCalculated){
			checkFullSubject();
		}
		return (percentage > 99.0);
	}

	public String getName() {
		return name;
	}

	public Map<Questionnaire, Double> getRelevantQuestionaires() {
		if(!avgCalculated){
			checkFullSubject();
		}
		return relevantQuestionaires;
	}

	public double getPercentage() {
		return percentage;
	}

	public int getFinalGrade() {
		if(!avgCalculated){
			checkFullSubject();
		}
		return finalGrade;
	}

	public void setNumOfUnits(int numOfUnits) {
		this.numOfUnits = numOfUnits;
	}

	public boolean hasPassed() {
		return getFinalGrade() >= Student.PASSING_BAR;
	}

	public boolean isMandatory() {
		return isMandatory;
	}

	public void setMandatory(boolean isMandatory) {
		this.isMandatory = isMandatory;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((name == null) ? 0 : name.hashCode());
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
		Subject other = (Subject) obj;
		if (name == null) {
			if (other.name != null)
				return false;
		} else if (!name.equals(other.name))
			return false;
		return true;
	}

	@Override
	public String toString() {
		return "Subject [name=" + name + ", numOfunits=" + numOfUnits + ", finalGrade=" + finalGrade +  ", questionnaires=" + 
					questionnaires + "]";
	}

	public Subject copy() {
		return new Subject(name, isMandatory, numOfUnits);
	}

	@Override
	public int compareTo(Subject other) {
		if(this.isMandatory){
			if(other.isMandatory){
				if(this.numOfUnits <= 2){
					if(other.numOfUnits <= 2){
						if(this.name.equals(other.name)){
							return this.numOfUnits - other.numOfUnits;
						}
						return this.name.compareTo(other.name);
					}
					return -1;
				}
				if(other.numOfUnits <= 2){
					return 1;
				}
				if(this.name.equals(other.name)){
					return this.numOfUnits - other.numOfUnits;
				}
				return this.name.compareTo(other.name);
			}
			return -1;
		}
		if(other.isMandatory){
			return 1;
		}
		if(this.name.equals(other.name)){
			return other.numOfUnits - this.numOfUnits;
		}
		return this.name.compareTo(other.name);		
	}

}
