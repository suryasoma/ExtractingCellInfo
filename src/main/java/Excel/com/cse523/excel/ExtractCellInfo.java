package Excel.com.cse523.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;




public class ExtractCellInfo {
	public static void main(String[] args) throws IOException {
//        String excelFilePath = "michelle_lokay__26590__IGSUpdate.xls";
        String excelFilePath = "errol_mclaughlin_jr__10717__Zipper LTD Reconciliation 10-01.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        
        Workbook workbook = WorkbookFactory.create(new File(excelFilePath));
        int numberOfSheets = workbook.getNumberOfSheets();
        int rowIndex = 0;
        for(int i=0; i<numberOfSheets; i++) {
        	Sheet sheet = workbook.getSheetAt(i);
        	Iterator<Row> iterator = sheet.iterator();
            
            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
                 
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if(sheet.getSheetName().equals("Annotation_Status_Data") && cell.getAddress().toString().equals("A4")) {
//                    	System.out.println("in F25");
                    }
                    if(cell.getCellType() == CellType.BLANK || cell.getCellType() == CellType._NONE || cell.getCellType() == CellType.ERROR) {
                    	continue;
                    }
                    String requiredInfo = "";
                    rowIndex++;
                    requiredInfo += rowIndex + ",";
                    requiredInfo += excelFilePath + ",";
                    requiredInfo += sheet.getSheetName() + ",";
                    requiredInfo += cell.getAddress().toString() + ",";
                    requiredInfo += ExtractUtil.isNumeric(cell) + ",";//cell type for numeric
                    requiredInfo += ExtractUtil.isString(cell) + ",";//cell type 1 for string
                    requiredInfo += ExtractUtil.lengthOfCell(cell) + ",";//length of words
                    requiredInfo += ExtractUtil.numberOfWords(cell) + ",";//number of words
                    requiredInfo += ExtractUtil.countLeadingSpaces(cell) + ",";//number of leading spaces
                    requiredInfo += ExtractUtil.isForstCharNum(cell) + ",";//is first character num
                    requiredInfo += ExtractUtil.isFirstCharSpecial(cell) + ",";//is first character special
                    requiredInfo += ExtractUtil.areWordsCapitalized(cell) + ",";//are words capitalized
                    requiredInfo += ExtractUtil.haveOnlyUpperCasedLetters(cell) + ",";//have only upper cased letters
                    requiredInfo += ExtractUtil.haveAlphaNumericCharacters(cell) + ",";//have alpha numeric characters
                    requiredInfo += ExtractUtil.haveAnySpecialCharacters(cell) + ",";//have special characters
                    requiredInfo += ExtractUtil.hasPunctuation(cell) + ",";//has punctuation
                    requiredInfo += ExtractUtil.hasColon(cell) + ",";//has colon
                    requiredInfo += ExtractUtil.hasWordTotal(cell) + ",";//has word total
                    requiredInfo += ExtractUtil.hasWordTable(cell) + ",";//has word table
                    
                    System.out.println(requiredInfo);

                }
            }
        }
        
         
        workbook.close();
        inputStream.close();
    }
}
