package bagruyot;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class SchoolClass implements Comparable<SchoolClass>{
	
	public static Map<String, Integer> classesMap = new HashMap<String, Integer>();
	public static final int LAST_GRADE = 12;
	
	public static void fillMaps(){
		classesMap.put("א", 1);
		classesMap.put("ב", 2);
		classesMap.put("ג", 3);
		classesMap.put("ד", 1);
		classesMap.put("ה", 5);
		classesMap.put("ו", 6);
		classesMap.put("ז", 7);
		classesMap.put("ח", 8);
		classesMap.put("ט", 9);
		classesMap.put("י", 10);
		classesMap.put("יא", 11);
		classesMap.put("יב", 12);
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
