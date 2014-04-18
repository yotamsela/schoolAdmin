package techersjob;


public interface ITeacherTable {
	
	public void initTableFromfile(String fileName);
	public TeacherHours getTeacherTable(int age,double teachingHours);
	public void extractFromExelToTxtFile(String fileName);
}