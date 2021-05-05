package Excel.com.cse523.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;




public class ExtractCellInfo {
	public static void main(String[] args) throws IOException {

		String csvFilePath = args[args.length-1];
        FileWriter csvWriter = new FileWriter(csvFilePath);
        
        String requiredInfo = "";
		requiredInfo += "Index,";
		requiredInfo += "File,";
		requiredInfo += "Sheet,";
		requiredInfo += "cell_address,";
		requiredInfo += "isNumeric,";
		requiredInfo += "isFormula,";
		requiredInfo += "length_of_cell,";
		requiredInfo += "number_of_words,";
		requiredInfo += "leading_spaces,";
		requiredInfo += "first_character_number,";
		requiredInfo += "first_character_special,";
		requiredInfo += "capitalized,";
		requiredInfo += "upper_cased,";
		requiredInfo += "contain_alpha_numeric,";
		requiredInfo += "contain_special_characters,";
		requiredInfo += "contain_punctuation,";
		requiredInfo += "contain_colon,";
		requiredInfo += "contain_total,";
		requiredInfo += "contain_table,";
		requiredInfo += "in_year_range,";
		requiredInfo += "aggr_formula_used,";
		requiredInfo += "reference_formula_value_numeric,";
		requiredInfo += "first_row_number,";
		requiredInfo += "first_column_number,";
		requiredInfo += "neighbors=0,";
		requiredInfo += "neighbors=1,";
		requiredInfo += "neighbors=2,";
		requiredInfo += "neighbors=3,";
		requiredInfo += "neighbors=4,";
		requiredInfo += "h_alignment=left,";
		requiredInfo += "h_alignment=center,";
		requiredInfo += "h_alignment=right,";
		requiredInfo += "h_alignment=default,";
		requiredInfo += "v_alignment=top,";
		requiredInfo += "v_alignment=center,";
		requiredInfo += "v_alignment=bottom,";
		requiredInfo += "indendation,";
		requiredInfo += "fill_pattern=default,";
		requiredInfo += "is_text_wrapped,";
		requiredInfo += "cell_size,";
		requiredInfo += "no_top_border,";
		requiredInfo += "thin_top_border,";
		requiredInfo += "no_bottom_border,";
		requiredInfo += "no_left_border,";
		requiredInfo += "no_right_border,";
		requiredInfo += "medium_right_border,";
		requiredInfo += "no_of_borders,";
		requiredInfo += "font_size,";
		requiredInfo += "is_default_font_color,";
		requiredInfo += "is_bold,";
		requiredInfo += "is_single_underlined,";
		requiredInfo += "matches_top_type,";
		requiredInfo += "matches_bottom_type,";
		requiredInfo += "matches_left_type,";
		requiredInfo += "matches_right_type,";
		requiredInfo += "matches_top_style,";
		requiredInfo += "matches_bottom_style,";
		requiredInfo += "matches_left_style,";
		requiredInfo += "matches_right_style,";
		requiredInfo += "top_neighbor_type,";
		requiredInfo += "bottom_neighbor_type,";
		requiredInfo += "left_neighbor_type,";
		requiredInfo += "right_neighbor_type";
		//System.out.println(requiredInfo);
		csvWriter.append(requiredInfo);
		csvWriter.append("\n");
		csvWriter.flush();

        int rowIndex = 0;
        
		for(int argsIndex=0; argsIndex<args.length-1; argsIndex++) {
			try {
				String fileName = args[argsIndex];
				String excelFilePath = "files/" + fileName;
		        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		        
		        Workbook workbook = WorkbookFactory.create(new File(excelFilePath));
		        int numberOfSheets = workbook.getNumberOfSheets();
	
				FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
				workbook.setForceFormulaRecalculation(true);
				
				
		        for(int i=0; i<numberOfSheets; i++) {
		        	Sheet sheet = workbook.getSheetAt(i);
		        	Iterator<Row> iterator = sheet.iterator();
		        	Set<CellAddress> mergedCells =  new HashSet<CellAddress>(); 
		        	HashMap<CellAddress, Integer> mergedCellsSize = new HashMap<CellAddress, Integer>();
		        	if(sheet.getSheetName().equals("Range_Annotations_Data") || sheet.getSheetName().equals("Annotation_Status_Data")) 
		        		continue;
		            
		            for(int j = 0; j < sheet.getNumMergedRegions(); ++j) {
		                CellRangeAddress range = sheet.getMergedRegion(j);
		                Iterator mergedIt = range.iterator();
		                boolean skipFirst = true;
		                while(mergedIt.hasNext()) {
		                	CellAddress cellAddr = (CellAddress)mergedIt.next();
		                	if(!skipFirst)
		                		mergedCells.add(cellAddr);
		                	else {
		                		mergedCellsSize.put(cellAddr, range.getNumberOfCells());
		                	}
		                	skipFirst = false;
		                	
		                }
		            }
		        	
		            while (iterator.hasNext()) {
		                Row nextRow = iterator.next();
		                Iterator<Cell> cellIterator = nextRow.cellIterator();
		                 
		                while (cellIterator.hasNext()) {
		                    Cell cell = cellIterator.next();
							
							if(cell.getSheet().getSheetName().toString().equals("Sept Liq Rec") &&
							 cell.getAddress().toString().equals("G23")) { 
								 System.out.println("hello"); 
							}
							 
		                    if(cell.getCellType() == CellType.BLANK || cell.getCellType() == CellType._NONE || cell.getCellType() == CellType.ERROR) {
		                    	continue;
		                    }
		                    if(mergedCells.contains(cell.getAddress()))
		                    	continue;
		                    requiredInfo = "";
		                    rowIndex++;
		                    requiredInfo += rowIndex + ",";
		                    requiredInfo += fileName + ",";
		                    requiredInfo += sheet.getSheetName() + ",";
		                    requiredInfo += cell.getAddress().toString() + ",";
		                    requiredInfo += ExtractUtil.isNumeric(cell) + ",";//cell type for numeric
		                    requiredInfo += ExtractUtil.isFormula(cell) + ",";//cell type 1 for string
		                    requiredInfo += ExtractUtil.lengthOfCell(cell) + ",";//length of words
		                    requiredInfo += ExtractUtil.numberOfWords(cell) + ",";//number of words
		                    requiredInfo += ExtractUtil.countLeadingSpaces(cell) + ",";//number of leading spaces
		                    requiredInfo += ExtractUtil.isFirstCharNum(cell) + ",";//is first character num
		                    requiredInfo += ExtractUtil.isFirstCharSpecial(cell) + ",";//is first character special
		                    requiredInfo += ExtractUtil.areWordsCapitalized(cell) + ",";//are words capitalized
		                    requiredInfo += ExtractUtil.haveOnlyUpperCasedLetters(cell) + ",";//have only upper cased letters
		                    requiredInfo += ExtractUtil.haveAlphaNumericCharacters(cell) + ",";//have alpha numeric characters
		                    requiredInfo += ExtractUtil.haveAnySpecialCharacters(cell) + ",";//have special characters
		                    requiredInfo += ExtractUtil.hasPunctuation(cell) + ",";//has punctuation
		                    requiredInfo += ExtractUtil.hasColon(cell) + ",";//has colon
		                    requiredInfo += ExtractUtil.hasWordTotal(cell) + ",";//has word total
		                    requiredInfo += ExtractUtil.hasWordTable(cell) + ",";//has word total
		                    requiredInfo += ExtractUtil.inYearRange(cell) + ",";//is in year range
		                    requiredInfo += ExtractUtil.isAggregateFormulaUsed(cell) + ",";//is aggregate formula used
		                    requiredInfo += ExtractUtil.isReferenceFormulaValueNumeric(formulaEvaluator, cell) + ",";//is reference formula value type numeric
		                    requiredInfo += ExtractUtil.firstRowNumer(cell) + ",";//first row number
		                    requiredInfo += ExtractUtil.firstColumnNumber(cell) + ",";//first column number
		                    requiredInfo += NeighbourUtil.numberOfNeighbours(sheet, cell) + ",";//number of neighbors
		                    requiredInfo += ExtractUtil.getHorizontalAlignment(cell) + ",";//get horizontal alignment
		                    requiredInfo += ExtractUtil.getVericalAlignment(cell) + ",";//get vertical alignment
		                    requiredInfo += ExtractUtil.getIndentation(cell) + ",";//get vertical alignment
		                    requiredInfo += ExtractUtil.isDefaultFillPattern(cell) + ",";//get vertical alignment
		                    requiredInfo += ExtractUtil.isTextWrapped(cell) + ",";//get vertical alignment
		                    requiredInfo += ExtractUtil.getCellSize(cell, mergedCellsSize) + ",";//get cell size
		                    requiredInfo += ExtractUtil.hasNoTopBorder(cell) + ",";//has no top border
		                    requiredInfo += ExtractUtil.hasThinTopBorder(cell) + ",";//has thin top border
		                    requiredInfo += ExtractUtil.hasNoBottomBorder(cell) + ",";//has no bottom border
		                    requiredInfo += ExtractUtil.hasNoLeftBorder(cell) + ",";//has no left border
		                    requiredInfo += ExtractUtil.hasNoRightBorder(cell) + ",";//has no right border
		                    requiredInfo += ExtractUtil.hasMediumRightBorder(cell) + ",";//has medium right border
		                    requiredInfo += ExtractUtil.getNumberOfBorders(cell) + ",";//number of borders
		                    requiredInfo += ExtractUtil.getFontSize(cell, workbook) + ",";//font height
		                    requiredInfo += ExtractUtil.isFontColorDefault(cell, workbook) + ",";//font color default
		                    requiredInfo += ExtractUtil.isBold(cell, workbook) + ",";//is bold
		                    requiredInfo += ExtractUtil.isSingleUnderlined(cell, workbook)+ ",";//id underlined
							requiredInfo += NeighbourUtil.checkNeighbourType(sheet, cell)+","; //get if neighbour type matches
							requiredInfo += NeighbourUtil.checkNeighbourStyle(sheet, cell) + ","; //get if neighbour type matches
							requiredInfo += NeighbourUtil.getNeighbourType(sheet, cell);

		                    //System.out.println(requiredInfo);
		            		csvWriter.append(requiredInfo);
		            		csvWriter.append("\n");
		            		csvWriter.flush();
	
		                }
		            }
		        }
		        
		        workbook.close();
		        inputStream.close();
			} catch(FileNotFoundException fe) {
				fe.printStackTrace();
			}
		} 
        csvWriter.close();
        
    }
}
