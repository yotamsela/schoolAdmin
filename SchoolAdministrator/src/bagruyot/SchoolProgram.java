package bagruyot;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import javax.swing.JOptionPane;
import org.eclipse.swt.widgets.Display;

public class SchoolProgram {
	
	private static boolean run;
	
	public static void main(String[] args) {
		if(testDate()){
			if(!ExamsReader.fillMaps()){
				return;
			}
			runShell();
		}
		else{
			JOptionPane.showMessageDialog(null, "תם תוקף השימוש בתוכנה\n אנא פנה לכתובת האי מייל:\n moshenfeld@gmail.com", "error", JOptionPane.ERROR_MESSAGE);			
		}
	}

	private static boolean testDate(){
		int year = 2014;
		int month = 05;
		int day = 01;
		String dateText = year + "/" + month + "/" + day;
		Date utilDate = null;
		try {
			SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM/dd");
			utilDate = formatter.parse(dateText);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		Date date = new Date();
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cal.setTime(utilDate);
		return date.before(utilDate);
	}

	private static void runShell() {
		try {
			Display display = Display.getDefault();
			MainShell shell = new MainShell(display);
			shell.open();
			shell.layout();
			run = true;
			while (run && !shell.isDisposed()) {
				if (!display.readAndDispatch()) {
					display.sleep();
				}
			}
			display.dispose();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void fullCalculation(boolean writeGrade10, boolean writeGrade11, boolean writeGrade12, boolean writeTeachers, 
			boolean writeStudents, boolean writeErrors, String filePath){
//		Model model = new Model(writeGrade10, writeGrade11, writeGrade12, writeTeachers, writeStudents, writeErrors, filePath);
		run = false;
		Model.fullCalculation(writeGrade10, writeGrade11, writeGrade12, writeTeachers, writeStudents, writeErrors, filePath);
//		new Thread(model).start();
	}
}
