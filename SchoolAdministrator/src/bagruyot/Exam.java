package bagruyot;
public class Exam {

	public static final int DISQUALIFIED_CODE = 777;
	public static final int EMPTY_EXAM_CODE = 666;

	private int magenGrade;
	private int examGrade;
	private int finalGrade;
	private int date;
	
	public Exam(int magenGrade, int examGrade, int finalGrade, int date) {
		this.magenGrade = magenGrade;
		this.examGrade = examGrade;
		this.finalGrade = finalGrade;
		if(examGrade == DISQUALIFIED_CODE || examGrade == EMPTY_EXAM_CODE){
			this.examGrade = 0;
			this.finalGrade = 0;
		}
		this.date = date;
	}

	public int getMagenGrade() {
		return magenGrade;
	}

	public int getExamGrade() {
		return examGrade;
	}

	public int getFinalGrade() {
		return finalGrade;
	}

	public int getYear() {
		return date / 100;
	}
	
	public String getDate() {
		return (((date % 100) < 10)? "0": "") + (date % 100) + "/" + (date / 100);
	}

	@Override
	public String toString() {
		return "Exam [magenGrade=" + magenGrade + ", examGrade=" + examGrade + ", finalGrade=" + finalGrade + ", date=" + date + "]";
	}
	
}
