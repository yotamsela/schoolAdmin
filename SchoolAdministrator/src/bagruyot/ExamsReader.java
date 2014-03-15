package bagruyot;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.jar.JarEntry;
import java.util.jar.JarFile;

import javax.swing.JOptionPane;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExamsReader {

	public static final String YES = "V";

	public static boolean fillMaps() {
		//TODO the user should choose?
		//		URL url = null;
		//		File file = null;
		//		try {
		//			url = new URL("jar:file:/" + new File("").toURI().toURL() + "schoolAdministrator.jar!/questionaires.xlsx");
		//			file = new File(url.toURI());
		//		} catch (MalformedURLException e) {
		//			JOptionPane.showMessageDialog(null, "1\n" + url, "error", JOptionPane.ERROR_MESSAGE);
		//			// TODO Auto-generated catch block
		//			e.printStackTrace();
		//		} catch (URISyntaxException e) {
		//			JOptionPane.showMessageDialog(null, "2\n" + url, "error", JOptionPane.ERROR_MESSAGE);
		//			// TODO Auto-generated catch block
		//			e.printStackTrace();
		//		}

		//		ExamsReader.class.getResourceAsStream("/questionaires.xlsx");

		//		File file = new File("questionaires.xlsx");

		//		ZipFile file = new ZipFile("whatever.jar");
		//		if (file != null) {
		//		   ZipEntries entries = file.entries(); //get entries from the zip file...
		//
		//		   if (entries != null) {
		//		      while (entries.hasMoreElements()) {
		//		          ZipEntry entry = entries.nextElement();
		//
		//		          //use the entry to see if it's the file '1.txt'
		//		          //Read from the byte using file.getInputStream(entry)
		//		      }
		//		    }
		//		}

		//		String inputPath = "path/to/.jar";
		//		String outputPath = "Desktop/CopyText.txt";
		//		File resStreamOut = new File(outputPath);
		//		int readBytes;
		//		JarFile file = new JarFile(inputPath);
		//		FileWriter fw = new FileWriter(resStreamOut);
		//		try{
		//			Enumeration<JarEntry> entries = file.entries();
		//			while (entries.hasMoreElements()){
		//				JarEntry entry = entries.nextElement();
		//				if (entry.getName().equals("readMe/tempReadme.txt")) {
		//
		//					System.out.println(entry +" : Entry");
		//					InputStream is = file.getInputStream(entry);
		//					BufferedWriter output = new BufferedWriter(fw);
		//					while ((readBytes = is.read()) != -1) {
		//						output.write((char) readBytes);
		//					}
		//					System.out.println(outputPath);
		//					output.close();
		//				} 
		//			}
		//		} catch(Exception er){
		//			er.printStackTrace();
		//		}

		InputStream is = ExamsReader.class.getClassLoader().getResourceAsStream("questionaires.xlsx");
		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(is);		
		XSSFSheet sheet = workbook.getSheetAt(0);
		int i = 1, size = sheet.getLastRowNum(), j, rowSize, qNumber, numOfUnits;
		double bonus, sum;
		XSSFRow firstRow, secondRow;
		Subject subject;
		Questionnaire questionnaire;
		List<Map<Questionnaire, Double>> maps;
		Map<Questionnaire, Double> map;
		while(i <= size){
			firstRow = sheet.getRow(i);
			subject = new Subject(firstRow.getCell(1).toString(), firstRow.getCell(4).toString().equals(YES), 0);
			numOfUnits = (int) firstRow.getCell(2).getNumericCellValue();
			bonus = firstRow.getCell(3).getNumericCellValue();
			Model.bonusesMap.put(subject, bonus);
			maps = new ArrayList<Map<Questionnaire, Double>>();
			i++;
			firstRow = sheet.getRow(i);
			secondRow = sheet.getRow(i + 1);
			while(firstRow != null && firstRow.getCell(1) != null && (firstRow.getCell(1).toString().length() > 0)){
				map = new HashMap<Questionnaire, Double>();
				map.put(null, (double)numOfUnits);
				rowSize = firstRow.getLastCellNum();
				sum = 0;
				for (j = 1; j < rowSize; j++) {
					qNumber = 0;
					if(firstRow.getCell(j) != null){
						qNumber = (int) firstRow.getCell(j).getNumericCellValue();
						questionnaire = Model.questionnaireMap.get(qNumber);
						if(Model.questionnaireMap.get(qNumber) == null){
							questionnaire = new Questionnaire(qNumber);
							Model.questionnaireMap.put(qNumber, questionnaire);							
							Model.subjectsMap.put(questionnaire, subject);
						}
						sum += secondRow.getCell(j).getNumericCellValue();
						map.put(questionnaire, secondRow.getCell(j).getNumericCellValue());
					}
				}
				if(sum < 99.0 || sum > 101.0){
					JOptionPane.showMessageDialog(null, "תקלה: סכום המשקלים של המקצוע - " + subject.getName() + " שווה ל-" + sum, "שגיאת שקלול אחוזים",
							JOptionPane.ERROR_MESSAGE);
					return false;
				}
				maps.add(map);
				i+=2;
				firstRow = sheet.getRow(i);
				secondRow = sheet.getRow(i + 1);
			}
			if(Model.percentagesMap.get(subject) != null){
				Model.percentagesMap.get(subject).addAll(maps);
			}
			else{
				Model.percentagesMap.put(subject, maps);				
			}
			i+=2;
		}
		return true;
	}

	//TODO does not work yet
	public static void outputFile() {
		//TODO chack the name instead
		String inputPath = "school administrator.jar";
		String outputPath = "שאלונים.xlsx";
		File resStreamOut = new File(outputPath);
		int readBytes;
		try{
			JarFile file = new JarFile(inputPath);
			File outFile = new File(outputPath);
			if(!outFile.exists()){
				outFile.createNewFile();
			}
			FileWriter fw = new FileWriter(resStreamOut);
			Enumeration<JarEntry> entries = file.entries();
			while (entries.hasMoreElements()){
				JarEntry entry = entries.nextElement();
				if (entry.getName().equals("questionaires.xlsx")) {
					InputStream is = file.getInputStream(entry);
					BufferedWriter output = new BufferedWriter(fw);
					while ((readBytes = is.read()) != -1) {
						output.write(readBytes);
					}
					System.out.println(outputPath);
					output.close();
				} 
			}
			file.close();
			fw.close();
		} catch(Exception er){
			er.printStackTrace();
		}
	}

}
