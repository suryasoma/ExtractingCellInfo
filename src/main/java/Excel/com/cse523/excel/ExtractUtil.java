package Excel.com.cse523.excel;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.ss.formula.Formula;
import org.apache.commons.lang3.StringUtils;

public class ExtractUtil {
	
	protected static int isNumeric(Cell cell) {
		if (cell.getCellType() == CellType.NUMERIC) {
			if(HSSFDateUtil.isCellDateFormatted(cell)) {
				return 0;
			}
			return 1;
		}
		return 0;
	}
	
	protected static int isFormula(Cell cell) {
		if (cell.getCellType() == CellType.FORMULA) {
			return 1;
		}
		return 0;
	}
	
	protected static int lengthOfCell(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue().length();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
//                return String.valueOf(cell.getDateCellValue()).length();
                DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
                Date date = cell.getDateCellValue();
                String cellValue = df.format(date);
                return cellValue.length();
            } else {
            	return String.valueOf(cell.getNumericCellValue()).length();
            }
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue()).length();
		case FORMULA:
//			EvaluationCell evaluationCell = evalSheet.getCell(cell.getRowIndex(), cell.getColumnIndex());
//            Ptg[] formulaTokens = evalWorkbook.getFormulaTokens(evaluationCell);
			return String.valueOf(cell.getCellFormula()).length();
		default:
			return 0;
		}
	}
	
	protected static int numberOfWords(Cell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
		case FORMULA:
			return 0;
		case STRING:
			return cell.getStringCellValue().trim().split("\\s").length;
		default:
			return 1;
		}
	}
	
	protected static int countLeadingSpaces(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
                cellValue = String.valueOf(cell.getDateCellValue());
            } else {
            	cellValue = String.valueOf(cell.getNumericCellValue());
            }
			break;
		case BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case FORMULA:
			cellValue = String.valueOf(cell.getCellFormula());
			break;
		default:
			break;
		}
		
		char[] chars = cellValue.toCharArray();
		int count = 0;
		for(int i=0; i<chars.length; i++) {
			if(chars[i] == ' ') {
				count++;
			} else {
				break;
			}
		}
		return count;
	}
	
	protected static int isFirstCharNum(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case NUMERIC:
			return 1;
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case FORMULA:
			cellValue = cell.getCellFormula();
			break;
		default:
			return 0;
		}
		char[] chars = cellValue.toCharArray();
		
		for(int i=0; i<chars.length; i++) {
			if(chars[i] != ' ') {
				if (chars[i] >= 48 && chars[i] <= 57) {
					return 1;
				} else {
					return 0;
				}
			}
		}
		return 0; 
	}
	
	
	protected static int isFirstCharSpecial(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case FORMULA:
//			cellValue = String.valueOf(cell.getCellFormula());
//			break;
		case NUMERIC:
		case BOOLEAN:
		default:
			return 0;
		}
//		cellValue = cellValue.trim();
		if(cellValue.length() == 0)
			return 0;
		String word = String.valueOf(cellValue.charAt(0));
		String regex = "[^A-Za-z0-9 ]*";
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(word);
		if(m.matches())
			return 1;
		else
			return 0;
	}
	
	protected static int areWordsCapitalized(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case NUMERIC:
		case FORMULA:
		default:
			return 0;
		}
		String[] words = cellValue.split("\\s+");
		if(words.length == 0)
			return 0;
		for(int i=0; i<words.length; i++) {
			char[] chars = words[i].toCharArray();
			if(chars.length > 0 && !(chars[0] >= 65 && chars[0] <= 97)){
				return 0;
			} 
		}
		return 1;
	}
	
	protected static int haveOnlyUpperCasedLetters(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue());
			break;
		case NUMERIC:
		case FORMULA:
		default:
			return 0;
		}
		String regex = "[^a-z]*";
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(cellValue);
		if(m.matches())
			return 1;
		else
			return 0;
		
		
		/*
		 * String[] words = cellValue.split("\\s+"); if(words.length == 0) return 0;
		 * 
		 * Pattern p = Pattern.compile(regex); for (String word : words) { Matcher m =
		 * p.matcher(word); if(!m.matches()) return 0; } return 1;
		 */
	}
	
	protected static int haveAlphaNumericCharacters(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case FORMULA:
			cellValue = cell.getCellFormula();
			break;
		case BOOLEAN:
			return 1;
		case NUMERIC:
		default:
			return 0;
		}
	return StringUtils.isAlphanumeric(cellValue) ? 1 : 0;
	}
	
	protected static int haveAnySpecialCharacters(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case FORMULA:
			cellValue = cell.getCellFormula();
			break;
		case BOOLEAN:
		case NUMERIC:
		default:
			return 0;
		}
		
		Pattern p = Pattern.compile("[^a-z0-9 ]", Pattern.CASE_INSENSITIVE);
		Matcher m = p.matcher(cellValue);
		if(m.find())
			return 1;
		else
			return 0;
		/*
		 * String[] words = cellValue.split("\\s+"); if(words.length == 0) return 0;
		 * 
		 * for (String word : words) { if(ExtractUtil.haveAnySpecialCharacters(word) ==
		 * 0) return 0; } return 1;
		 */
	}
	
	protected static int hasColon(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case FORMULA:
			cellValue = cell.getCellFormula();
			break;
		case BOOLEAN:
		case NUMERIC:
		default:
			return 0;
		}
		
		int index = cellValue.indexOf(":");
		if(index >= 0)
			return 1;
		else
			return 0;
	}
	
	protected static int hasPunctuation(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case FORMULA:
			cellValue = cell.getCellFormula();
			break;
		case BOOLEAN:
		case NUMERIC:
		default:
			return 0;
		}
		
		Pattern p = Pattern.compile("[.,;!?\"()]");
		Matcher m = p.matcher(cellValue);
		if(m.find()) {
			return 1;
		}
		else 
			return 0;
	}
	
	protected static int hasWordTotal(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue().toLowerCase();
			break;
		case FORMULA:
		case BOOLEAN:
		case NUMERIC:
		default:
			return 0;
		}
		
		String regex = ".*\\b" + Pattern.quote("total") + "\\b.*"; // \b is a word boundary
		if(cellValue.toLowerCase().matches(regex)) {
			return 1;
		} else {
			return 0;
		}
	}
	
	protected static int hasWordTable(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue().toLowerCase();
			break;
		case FORMULA:
		case BOOLEAN:
		case NUMERIC:
		default:
			return 0;
		}
		
		String regex = ".*\\b" + Pattern.quote("table") + "\\b.*"; // \b is a word boundary
		if(cellValue.toLowerCase().matches(regex)) {
			return 1;
		} else {
			return 0;
		}
	}
	
	protected static int inYearRange(Cell cell) {
		switch (cell.getCellType()) {
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
                return 1;
            }
			Integer cellValue = Double.valueOf(cell.getNumericCellValue()).intValue();
			if(cellValue >= 1970 && cellValue <= 2099)
				return 1;
			else
				return 0;
		case STRING:
		case FORMULA:
		case BOOLEAN:
		default:
			return 0;
		}
	}
	
	protected static int isAggregateFormulaUsed(Cell cell) {
		if(cell.getCellType() == CellType.FORMULA && cell.getCellFormula().contains("SUM")) {
			return 1;
		}
		return 0;
	}
	
	protected static int firstRowNumer(Cell cell) {
		return cell.getRowIndex();
	}
	
	protected static int firstColumnNumber(Cell cell) {
		return cell.getColumnIndex();
	}
	
	protected static String getHorizontalAlignment(Cell cell) {
		String retValue = "0,0,0,0";
		CellStyle cellStyle = cell.getCellStyle();
		switch(cellStyle.getAlignment()) {
		case LEFT:
			retValue = "1,0,0,0";
			break;
		case CENTER:
			retValue = "0,1,0,0";
			break;
		case RIGHT:
			retValue = "0,0,1,0";
			break;
		case GENERAL:
			retValue = "0,0,0,1";
			break;
		}
		return retValue;
	}
	
	protected static String getVericalAlignment(Cell cell) {
		String retValue = "0,0,0";
		CellStyle cellStyle = cell.getCellStyle();
		switch(cellStyle.getVerticalAlignment()) {
		case TOP:
			retValue = "1,0,0";
			break;
		case CENTER:
			retValue = "0,1,0";
			break;
		case BOTTOM:
			retValue = "0,0,1";
			break;
		}
		return retValue;
	}
	
	protected static int getIndentation(Cell cell) {
		return cell.getCellStyle().getIndention();
	}
	
	protected static int isDefaultFillPattern(Cell cell) {
		return cell.getCellStyle().getFillPattern() == FillPatternType.NO_FILL ? 1 : 0;
	}
	
	protected static int isTextWrapped(Cell cell) {
		return cell.getCellStyle().getWrapText() ? 1 : 0;
	}
	
	protected static int getCellSize(Cell cell, HashMap<CellAddress, Integer> mergedCellsSize) {
		if(mergedCellsSize.containsKey(cell.getAddress())) {
			return mergedCellsSize.get(cell.getAddress());
		}
		return 1;
	}
	
	protected static int hasNoTopBorder(Cell cell) {
		return cell.getCellStyle().getBorderTop() == BorderStyle.NONE ? 1 : 0;
	}
	
	protected static int hasThinTopBorder(Cell cell) {
		return cell.getCellStyle().getBorderTop() == BorderStyle.THIN ? 1 : 0;
	}
	
	protected static int hasNoBottomBorder(Cell cell) {
		return cell.getCellStyle().getBorderBottom() == BorderStyle.NONE ? 1 : 0;
	}
	
	protected static int hasNoLeftBorder(Cell cell) {
		return cell.getCellStyle().getBorderLeft() == BorderStyle.NONE ? 1 : 0;
	}
	
	protected static int hasNoRightBorder(Cell cell) {
		return cell.getCellStyle().getBorderRight() == BorderStyle.NONE ? 1 : 0;
	}
	
	protected static int hasMediumRightBorder(Cell cell) {
		return cell.getCellStyle().getBorderRight() == BorderStyle.MEDIUM ? 1 : 0;
	}
	
	protected static int getNumberOfBorders(Cell cell) {
		int count = 0;
		if(cell.getCellStyle().getBorderTop() != BorderStyle.NONE)
			count+=1;
		if(cell.getCellStyle().getBorderBottom() != BorderStyle.NONE)
			count+=1;
		if(cell.getCellStyle().getBorderRight() != BorderStyle.NONE)
			count+=1;
		if(cell.getCellStyle().getBorderLeft() != BorderStyle.NONE)
			count+=1;
		return count;
	}
	
	protected static int getFontSize(Cell cell, Workbook wb) {
		CellStyle cellStyle = (CellStyle)cell.getCellStyle();
		if(cellStyle instanceof XSSFCellStyle)
			return ((XSSFCellStyle)cellStyle).getFont().getFontHeightInPoints();
		if(cellStyle instanceof HSSFCellStyle)
			return ((HSSFCellStyle)cellStyle).getFont(wb).getFontHeightInPoints();
		return 0;
	}
	
	protected static int isFontColorDefault(Cell cell, Workbook wb) {
		CellStyle cellStyle = (CellStyle)cell.getCellStyle();
		if(cellStyle instanceof XSSFCellStyle)
			return ((XSSFCellStyle)cellStyle).getFont().getColor() == XSSFFont.DEFAULT_FONT_COLOR ? 1 : 0;
		if(cellStyle instanceof HSSFCellStyle)
			return ((HSSFCellStyle)cellStyle).getFont(wb).getColor() == XSSFFont.DEFAULT_FONT_COLOR ? 1 : 0;
		return 0;
	}
	
	protected static int isBold(Cell cell, Workbook wb) {
		CellStyle cellStyle = (CellStyle)cell.getCellStyle();
		if(cellStyle instanceof XSSFCellStyle)
			return ((XSSFCellStyle)cellStyle).getFont().getBold() ? 1 : 0;
		if(cellStyle instanceof HSSFCellStyle)
			return ((HSSFCellStyle)cellStyle).getFont(wb).getBold() ? 1 : 0;
		return 0;
	}
	
	protected static int isSingleUnderlined(Cell cell, Workbook wb) {
		CellStyle cellStyle = (CellStyle)cell.getCellStyle();
		if(cellStyle instanceof XSSFCellStyle)
			return ((XSSFCellStyle)cellStyle).getFont().getUnderline() == Font.SS_NONE ? 1 : 0;
		if(cellStyle instanceof HSSFCellStyle)
			return ((HSSFCellStyle)cellStyle).getFont(wb).getUnderline() == Font.SS_NONE ? 1 : 0;
		return 0;
	}
	
	protected static String getStringWithTrimmedText(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue().trim();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
                DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
                Date date = cell.getDateCellValue();
                String cellValue = df.format(date);
                return cellValue.trim();
            } else {
            	return String.valueOf(cell.getNumericCellValue()).trim();
            }
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue()).trim();
		case FORMULA:
			return String.valueOf(cell.getCellFormula()).trim();
		default:
			return "";
		}
	}
	
	protected static int isReferenceFormulaValueNumeric(FormulaEvaluator formulaEvaluator, Cell cell) {
		if(cell.getCellType() == CellType.FORMULA) {
			try {
				return formulaEvaluator.evaluateFormulaCell(cell) == CellType.NUMERIC ? 1 : 0;
			} catch(Exception e) {
				return 0;
			}
		}
		return 0;
	}


}
