package bagruyot;


import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Iterator;

import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcellReaderAndWriter {

	public static XSSFWorkbook openXlsxFile(File file){
		File newFile = null;
		if(file.getName().endsWith(".xls")){
			System.out.println("i'm here");
			newFile = new File(file.getAbsolutePath() + "x");
			ExcellReaderAndWriter.convertToXlsx(file);
			file = newFile;
			newFile.deleteOnExit();
		}
		FileInputStream stream = null;
		try {
			stream = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(null, "failed to open file", "error", JOptionPane.ERROR_MESSAGE);
		}
		return openXlsxFile(stream);
	}

	public static XSSFWorkbook openXlsxFile(InputStream stream){
		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(stream);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, "failed to open file", "error", JOptionPane.ERROR_MESSAGE);
		}
		return workbook;
	}

	public static void convertToXlsx(File file){
		FileInputStream in = null;
		try {
			in = new FileInputStream(file);
		} catch (FileNotFoundException e) {
			JOptionPane.showMessageDialog(null, "failed to open file", "error", JOptionPane.ERROR_MESSAGE);
		}
		HSSFWorkbook oldWorkbook = null;
		try {
			oldWorkbook = new HSSFWorkbook(in);
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, "failed to open file", "error", JOptionPane.ERROR_MESSAGE);
		}
		XSSFWorkbook newWorkbook = new XSSFWorkbook();
		try {
			File outF = new File(file.getAbsolutePath() + "x");
			if (outF.exists())
				outF.delete();
			int sheetCnt = oldWorkbook.getNumberOfSheets();
			for (int i = 0; i < sheetCnt; i++) {
				Sheet sIn = oldWorkbook.getSheetAt(0);
				Sheet sOut;
				if(newWorkbook.getSheet(sIn.getSheetName()) == null){
					sOut = newWorkbook.createSheet(sIn.getSheetName()); 	
				}
				else{
					sOut = newWorkbook.getSheet(sIn.getSheetName());
				}
				Iterator<Row> rowIt = sIn.rowIterator();
				while (rowIt.hasNext()) {
					Row rowIn = rowIt.next();
					Row rowOut = sOut.createRow(rowIn.getRowNum());
					ExcellReaderAndWriter.copyRow(rowIn, rowOut);
				}
			}
			OutputStream out = new BufferedOutputStream(new FileOutputStream(outF));
			try {
				newWorkbook.write(out);
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} finally {
			try {
				in.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public static void copyRow(Row rowIn, Row rowOut){
		Iterator<Cell> cellIt = rowIn.cellIterator();
		while (cellIt.hasNext()) {
			Cell cellIn = cellIt.next();
			Cell cellOut = rowOut.createCell(
					cellIn.getColumnIndex(), cellIn.getCellType());
	
			switch (cellIn.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				break;
	
			case Cell.CELL_TYPE_BOOLEAN:
				cellOut.setCellValue(cellIn.getBooleanCellValue());
				break;
	
			case Cell.CELL_TYPE_ERROR:
				cellOut.setCellValue(cellIn.getErrorCellValue());
				break;
	
			case Cell.CELL_TYPE_FORMULA:
				cellOut.setCellFormula(cellIn.getCellFormula());
				break;
	
			case Cell.CELL_TYPE_NUMERIC:
				cellOut.setCellValue(cellIn.getNumericCellValue());
				break;
	
			case Cell.CELL_TYPE_STRING:
				cellOut.setCellValue(cellIn.getStringCellValue());
				break;
			}
	
			{
				CellStyle styleIn = cellIn.getCellStyle();
				CellStyle styleOut = cellOut.getCellStyle();
				styleOut.setDataFormat(styleIn.getDataFormat());
			}
			cellOut.setCellComment(cellIn.getCellComment());
		}
	}

}
