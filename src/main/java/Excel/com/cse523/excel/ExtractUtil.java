package Excel.com.cse523.excel;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class ExtractUtil {
	protected static int isNumeric(Cell cell) {
		if (cell.getCellType() == CellType.NUMERIC) {
			return 1;
		}
		return 0;
	}
	
	protected static int isString(Cell cell) {
		if (cell.getCellType() == CellType.STRING || cell.getCellType() == CellType.BOOLEAN) {
			return 1;
		}
		return 0;
	}
	
	protected static int lengthOfCell(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue().length();
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue()).length();
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue()).length();
		case FORMULA:
			return String.valueOf(cell.getCellFormula()).length();
		default:
			return 0;
		}
	}
	
	protected static int numberOfWords(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue().split("\\s+").length;
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
			cellValue = String.valueOf(cell.getNumericCellValue());
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
	
	protected static int isForstCharNum(Cell cell) {
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
	
	protected static int haveAnySpecialCharacters(String word) {
		String regex = "^[-+_!@#$%^&*.,?]+$";
		Pattern p = Pattern.compile(regex);
		Matcher m = p.matcher(word);
		if(m.matches())
			return 1;
		else
			return 0;
	}
	
	protected static int isFirstCharSpecial(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case STRING:
			cellValue = cell.getStringCellValue();
			break;
		case FORMULA:
			cellValue = String.valueOf(cell.getCellFormula());
			break;
		case NUMERIC:
		case BOOLEAN:
		default:
			return 0;
		}
		
		return ExtractUtil.haveAnySpecialCharacters(String.valueOf(cellValue.charAt(0)));
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
		String regex = "^[A-Z]+$";
		String[] words = cellValue.split("\\s+");
		if(words.length == 0)
			return 0;
		
		Pattern p = Pattern.compile(regex);
		for (String word : words) {
			Matcher m = p.matcher(word);
			if(!m.matches())
				return 0;
		}
		return 1;
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
		String regex = "^[a-zA-Z0-9]+$";
		String[] words = cellValue.split("\\s+");
		if(words.length == 0)
			return 0;
		
		Pattern p = Pattern.compile(regex);
		for (String word : words) {
			Matcher m = p.matcher(word);
			if(!m.matches())
				return 0;
		}
		return 1;
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
		String[] words = cellValue.split("\\s+");
		if(words.length == 0)
			return 0;
		
		for (String word : words) {
			if(ExtractUtil.haveAnySpecialCharacters(word) == 0)
				return 0;
		}
		return 1;
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
		
		String[] words = cellValue.split(":");
		if(words.length > 1)
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
		
		String[] words = cellValue.split("!");
		if(words.length > 1)
			return 1;
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
		
		String[] words = cellValue.split("total");
		if(words.length > 1)
			return 1;
		else
			return 0;
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
		
		String[] words = cellValue.split("table");
		if(words.length > 1)
			return 1;
		else
			return 0;
	}
}
