package techersjob;


public interface ITeacherTable {
	
	public void initTableFromfile(String fileName);
	public TeacherHours getTeacherTable(TeacherAttributes teacherAttributes);
	public void extractFromExelToTxtFile(String fileName);
}