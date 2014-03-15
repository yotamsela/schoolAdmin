package techersjob;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import bagruyot.ExcellReaderAndWriter;
import bagruyot.Writer;

public class Main {
	
	public static void main(String[] args) throws IOException {
		System.out.println("start");
		String pathname = "temp.xlsx";
		File file = new File(pathname);
//		OutputStream out = new BufferedOutputStream(new FileOutputStream(outF));
//		XSSFWorkbook newWorkbok = new XSSFWorkbook();
		
		XSSFWorkbook workbook = ExcellReaderAndWriter.openXlsxFile(file);
		XSSFSheet sheet = workbook.getSheetAt(1);
		double frontalSum = getFrontalSum(sheet);
		double gmolSum = getGmolSum(sheet);
		
		
		System.out.println("value: "+frontalSum);
		//xlsxFile.createSheet("hello2");
		//XSSFSheet sheet = workbook.createSheet("hello");
//		Writer.createCell(workbook, sheet.createRow(0), 0, "100", XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, XSSFCellStyle.BORDER_MEDIUM, 
//				XSSFCellStyle.BORDER_THIN, true, (short)11, true, false);
		
		Writer.write(workbook, "temp.xlsx");
		System.out.println("end");
	}
	
	private static double getFrontalSum(XSSFSheet sheet){
		XSSFRow row = sheet.getRow(24);
		XSSFCell cell = row.getCell(4);
		double value = cell.getNumericCellValue();
		System.out.println("value: "+value);
		return value;
	}
	
	private static double getGmolSum(XSSFSheet sheet){
		XSSFRow row = sheet.getRow(24);
		XSSFCell cell = row.getCell(7);
		double value = cell.getNumericCellValue();
		System.out.println("value: "+value);
		return value;
	}

}
