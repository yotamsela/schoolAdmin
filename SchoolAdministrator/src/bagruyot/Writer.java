package bagruyot;
import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Writer {

	private static Map<CellRepresentor, XSSFCellStyle> cellsMap = new HashMap<CellRepresentor, XSSFCellStyle>();

	public static void write(XSSFWorkbook workbook, String name) {
		//		boolean tryToPrint = true;
		File file = null;
		//		while(tryToPrint){
		try
		{
			file = new File(name);
			if(file.exists()){
				file.delete();
			}
			FileOutputStream out = new FileOutputStream(new File(name));
			workbook.write(out);
			out.close();
		}
		catch (Exception e)
		{//TODO fix it
			//				Object[] options = {"נסה שוב", "צא מהתוכנית"};
			//				int n = JOptionPane.showOptionDialog(null, "קיימת בעיה ביצירת הקובץ. יתכן שקובץ בשם \" " + file.getName() + 
			//						"\" כבר קיים והוא פתוח כעת. במידה וכן - יש לסגור את הקובץ ולנסות שוב", "בעיה ביצירת הקובץ", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE,
			//						null, options, options[1]);
			//				tryToPrint = n == 0;
		}
		//		}
	}
	
	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, int value, short upBorder, short downBorder,
			short leftBorder, short rightBorder, boolean bold, short fontSize){
		createCell(workbook, row, index, value, upBorder, downBorder, leftBorder, rightBorder, bold, fontSize, false, false);
	}

	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, double value, short upBorder, short downBorder,
			short leftBorder, short rightBorder, boolean bold, short fontSize, int numOfDigits){
		createCell(workbook, row, index, value, upBorder, downBorder, leftBorder, rightBorder, bold, fontSize, false, false, numOfDigits);
	}

	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, String value, short upBorder, short downBorder,short leftBorder, 
			short rightBorder, boolean bold, short fontSize, boolean wrapText, boolean rotate){
		createCell(workbook, row, index, value, upBorder, downBorder, leftBorder, rightBorder, fontSize, IndexedColors.BLACK.getIndex(), 
				IndexedColors.WHITE.getIndex(), bold, wrapText, rotate, false);
	}

	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, int value, short upBorder, short downBorder,
			short leftBorder, short rightBorder, boolean bold, short fontSize, boolean wrapText, boolean rotate){
		createCell(workbook, row, index, value, upBorder, downBorder, leftBorder, rightBorder, fontSize, IndexedColors.BLACK.getIndex(), 
				IndexedColors.WHITE.getIndex(), bold, wrapText, rotate, false);
	}

	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, double value, short upBorder, short downBorder,
			short leftBorder, short rightBorder, boolean bold, short fontSize, boolean wrapText, boolean rotate, int numOfDigits){
		createCell(workbook, row, index, value, upBorder, downBorder, leftBorder, rightBorder, fontSize, IndexedColors.BLACK.getIndex(), 
				IndexedColors.WHITE.getIndex(), bold, wrapText, rotate, false, numOfDigits);
	}

	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, String value, short upBorder, short downBorder, short leftBorder, 
			short rightBorder, short fontSize, short fontColor, short backgroundColor, boolean bold, boolean wrapText, boolean rotate, boolean underLine){
		CellRepresentor cr = new CellRepresentor(workbook, upBorder, downBorder, leftBorder, rightBorder, fontSize, fontColor, backgroundColor, bold, 
				rotate, underLine, wrapText, 0);
		XSSFCellStyle style = cellsMap.get(cr);
		if(style == null){
			style = createStyle(cr);
			cellsMap.put(cr, style);
		}		
		XSSFCell cell = row.createCell(index);
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}

	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, int value, short upBorder, short downBorder, short leftBorder, 
			short rightBorder, short fontSize, short fontColor, short backgroundColor, boolean bold, boolean wrapText, boolean rotate, boolean underLine){
		CellRepresentor cr = new CellRepresentor(workbook, upBorder, downBorder, leftBorder, rightBorder, fontSize, fontColor, 
				backgroundColor, bold, rotate, underLine, wrapText, 0);
		XSSFCellStyle style = cellsMap.get(cr);
		if(style == null){
			style = createStyle(cr);
			cellsMap.put(cr, style);
		}		
		XSSFCell cell = row.createCell(index);
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}

	public static void createCell(XSSFWorkbook workbook, XSSFRow row, int index, double value, short upBorder, short downBorder, short leftBorder, 
			short rightBorder, short fontSize, short fontColor, short backgroundColor, boolean bold, boolean wrapText, boolean rotate, boolean underLine,
			int numOfDigits){
		if((int) (value * Math.pow(10, numOfDigits)) == (value * Math.pow(10, numOfDigits))){
			numOfDigits = 0;
		}
		CellRepresentor cr = new CellRepresentor(workbook, upBorder, downBorder, leftBorder, rightBorder, fontSize, fontColor, 
				backgroundColor, bold, rotate, underLine, wrapText, numOfDigits);
		XSSFCellStyle style = cellsMap.get(cr);
		if(style == null){
			style = createStyle(cr);
			if(numOfDigits != 0){
				XSSFDataFormat df = workbook.createDataFormat();
				String format = "0.";
				for (int i = 0; i < numOfDigits; i++) {
					format = format + "#";
				}
				style.setDataFormat(df.getFormat(format));				
			}
			cellsMap.put(cr, style);
		}
		XSSFCell cell = row.createCell(index);
		if(numOfDigits != 0){
			cell.setCellValue(value);
		}
		else{
			cell.setCellValue((int)value);
		}
		cell.setCellStyle(style);
	}

	public static void changeCellStyle(XSSFWorkbook workbook, XSSFCell cell, short upBorder, short downBorder, short leftBorder, short rightBorder, 
			int numOfDigits){
		CellRepresentor cr = createCellRepresentor(workbook, cell, upBorder, downBorder, leftBorder, rightBorder, numOfDigits);
		XSSFCellStyle style = cellsMap.get(cr);
		if(style == null){
			style = createStyle(cr);
			cellsMap.put(cr, style);
		}
		cell.setCellStyle(style);
	}

	private static CellRepresentor createCellRepresentor(XSSFWorkbook workbook, XSSFCell cell, short upBorder,short downBorder, short leftBorder, 
			short rightBorder, int numOfDigits) {
		XSSFCellStyle style = cell.getCellStyle();
		XSSFFont font = style.getFont();
		short fontSize = font.getFontHeightInPoints();
		short fontColor = font.getColor();
		short backgroundColor = style.getFillForegroundColor();
		boolean bold = font.getBold();
		boolean wrapText = style.getWrapText();
		boolean rotate = style.getRotation() != 0;
		boolean underLine = font.getUnderline() == Font.U_SINGLE;
		return new CellRepresentor(workbook, upBorder, downBorder, leftBorder, rightBorder, fontSize, fontColor, backgroundColor, bold, 
				rotate, underLine, wrapText, numOfDigits);
	}

	private static XSSFCellStyle createStyle(CellRepresentor cr){
		XSSFCellStyle style = cr.workbook.createCellStyle();
		style.setBorderTop(cr.upBorder);
		style.setBorderBottom(cr.downBorder);
		style.setBorderLeft(cr.leftBorder);
		style.setBorderRight(cr.rightBorder);		
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFillForegroundColor(cr.backgroundColor);
		style.setWrapText(cr.wrapText);
		if(cr.rotate){
			style.setRotation((short)90);
		}
		XSSFFont font = cr.workbook.createFont();
		font.setFontHeightInPoints(cr.fontSize);
		font.setColor(cr.fontColor);
		font.setBold(cr.bold);
		if(cr.underLine){
			font.setUnderline(Font.U_SINGLE);				
		}
		style.setFont(font);
		return style;
	}

	private static class CellRepresentor {
		XSSFWorkbook workbook;
		short upBorder;
		short downBorder;
		short leftBorder;
		short rightBorder;
		short fontSize;
		short fontColor;
		short backgroundColor;
		boolean bold;
		boolean wrapText;
		boolean rotate;
		boolean underLine;
		int numOfDigits;

		private CellRepresentor(XSSFWorkbook workbook, short upBorder, short downBorder, short leftBorder, short rightBorder, short fontSize, 
				short fontColor, short backgroundColor, boolean bold, boolean rotate, boolean underLine, boolean wrapText, int numOfDigits) {
			this.workbook = workbook;
			this.upBorder = upBorder;
			this.downBorder = downBorder;
			this.leftBorder = leftBorder;
			this.rightBorder = rightBorder;
			this.fontSize = fontSize;
			this.fontColor = fontColor;
			this.backgroundColor = backgroundColor;
			this.bold = bold;
			this.rotate = rotate;
			this.underLine = underLine;
			this.wrapText = wrapText;
			this.numOfDigits = numOfDigits;
		}

		@Override
		public int hashCode() {
			final int prime = 31;
			int result = 1;
			result = prime * result + backgroundColor;
			result = prime * result + (bold ? 1231 : 1237);
			result = prime * result + downBorder;
			result = prime * result + fontColor;
			result = prime * result + fontSize;
			result = prime * result + leftBorder;
			result = prime * result + numOfDigits;
			result = prime * result + rightBorder;
			result = prime * result + (rotate ? 1231 : 1237);
			result = prime * result + (wrapText ? 1231 : 1237);
			result = prime * result + (underLine ? 1231 : 1237);
			result = prime * result + upBorder;
			result = prime * result
					+ ((workbook == null) ? 0 : workbook.hashCode());
			return result;
		}

		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (obj == null)
				return false;
			if (getClass() != obj.getClass())
				return false;
			CellRepresentor other = (CellRepresentor) obj;
			if (backgroundColor != other.backgroundColor)
				return false;
			if (bold != other.bold)
				return false;
			if (downBorder != other.downBorder)
				return false;
			if (fontColor != other.fontColor)
				return false;
			if (fontSize != other.fontSize)
				return false;
			if (leftBorder != other.leftBorder)
				return false;
			if (numOfDigits != other.numOfDigits)
				return false;
			if (rightBorder != other.rightBorder)
				return false;
			if (rotate != other.rotate)
				return false;
			if (wrapText != other.wrapText)
				return false;
			if (underLine != other.underLine)
				return false;
			if (upBorder != other.upBorder)
				return false;
			if (workbook == null) {
				if (other.workbook != null)
					return false;
			} else if (!workbook.equals(other.workbook))
				return false;
			return true;
		}
	}

}
