package bagruyot;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class SchoolClass implements Comparable<SchoolClass>{
	
	public static Map<String, Integer> classesMap = new HashMap<String, Integer>();
	public static final int LAST_GRADE = 12;
	
	public static void fillMaps(){
		classesMap.put("�", 1);
		classesMap.put("�", 2);
		classesMap.put("�", 3);
		classesMap.put("�", 1);
		classesMap.put("�", 5);
		classesMap.put("�", 6);
		classesMap.put("�", 7);
		classesMap.put("�", 8);
		classesMap.put("�", 9);
		classesMap.put("�", 10);
		classesMap.put("��", 11);
		classesMap.put("��", 12);
	}
	
	//-----------------------------------------------------------------------------
	
	private String name;
	private int year;	
	private List<Student> students;
	
	public SchoolClass(String name, int year){
		this.name = name;
		this.year = year;
		students = new ArrayList<Student>();
	}
	
	public int getYear() {
		return year;
	}

	public String getName(){
		return name;
	}

	@Override
	public int compareTo(SchoolClass other) {
		if(this.year != other.year){
			return other.year - this.year;
		}
		return this.name.compareTo(other.name);
	}
	
	public void addStudent(Student student){
		students.add(student);
	}

	public List<Student> getStudents() {
		return students;
	}
	
}
