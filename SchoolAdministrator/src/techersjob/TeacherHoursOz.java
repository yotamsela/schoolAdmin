package techersjob;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import techersjob.TabeleSeperator.TableType;
import bagruyot.ExcellReaderAndWriter;

public class TeacherHoursOz implements ITeacherTable{

	private Map<TabeleSeperator, Map<Double, TeacherHours>> map;
	private TableType type = TableType.OZ;

	private int[] startingRowLocationArray = {2,9,17,25,32};
	private int startingColumn = 1;
	private int size = 5;

	@Override
	public void initTableFromfile(
			String fileName) {
		String sCurrentLine;

		BufferedReader br;
		try {
			br = new BufferedReader(new FileReader(fileName));


			while ((sCurrentLine = br.readLine()) != null) {
				System.out.println(sCurrentLine);
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@Override
	public TeacherHours getTeacherTable(TeacherAttributes teacherAttributes) {

		if(!teacherAttributes.isOzTmura()){
			return TeacherHours.getZeroHours();
		}

		int age = teacherAttributes.getTeacherAge();
		boolean isMother = teacherAttributes.isChildOver14();
		TabeleSeperator range = TabeleSeperator.getRange(age, type, isMother);
		if(range == null){
			return null;
		}
		Map<Double, TeacherHours> innerMap = map.get(range);
		double teachingHoursKey = teacherAttributes.getGmol() + teacherAttributes.getFrontalHours();
		return innerMap.get(teachingHoursKey);
	}

	@Override
	public void extractFromExelToTxtFile(String fileName) {
		File file = new File(fileName);

		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		XSSFCell cell = null;
		try{
			PrintWriter writer = new PrintWriter(fileName.replaceAll("xlsx","txt"), "UTF-8");
			for (int k = 0; k < startingRowLocationArray.length; k++) {
				System.out.println("init location: "+startingRowLocationArray[k]);
				List<List<String>> rowList = new ArrayList<List<String>>();

				for (int i = this.startingRowLocationArray[k]; i < this.startingRowLocationArray[k] + size; i++) {
					XSSFRow row = sheet.getRow(i);
					int j = this.startingColumn;
					List<String> cellList = new ArrayList<String>();

					while((cell = row.getCell(j))!= null){
						j++;
						DecimalFormat df = new DecimalFormat("#.#");
						double numericCellValue = cell.getNumericCellValue();
						cellList.add(df.format(numericCellValue));

					}

					rowList.add(cellList);

				}

				for (int i = 0; i < rowList.get(0).size(); i++) {
					String line = "";
					for (int j = 0; j < rowList.size(); j++) {
						line += rowList.get(j).get(i) +" ";
					}
					System.out.println(line);
					writer.println(line);
				}
				writer.close();

			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}



}
