package Excel.com.cse523.excel;

import org.apache.poi.ss.usermodel.*;

public class NeighbourUtil {

    private static Cell getCellIfExists(Sheet sheet, int rowNumber, int columnIndex) {
        if(rowNumber < 0 || columnIndex < 0) {
            return null;
        }
        Row row = sheet.getRow(rowNumber);
        if(row == null)
            return null;
        Cell cell = row.getCell(columnIndex);
        if(cell == null || cell.getCellType() == CellType.BLANK)
            return null;
        return cell;
    }

    protected static String numberOfNeighbours(Sheet sheet, Cell cell) {
        int rowNumber = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();

        int count = 0;
        Cell leftCell = getCellIfExists(sheet, rowNumber, columnIndex-1);
        if(leftCell != null && ExtractUtil.getStringWithTrimmedText(leftCell).length() > 0)
            count++;

        Cell rightCell = getCellIfExists(sheet, rowNumber, columnIndex+1);
        if(rightCell != null && ExtractUtil.getStringWithTrimmedText(rightCell).length() > 0)
            count++;

        Cell topCell = getCellIfExists(sheet, rowNumber-1, columnIndex);
        if(topCell != null && ExtractUtil.getStringWithTrimmedText(topCell).length() > 0)
            count++;

        Cell bottomCell = getCellIfExists(sheet, rowNumber+1, columnIndex);
        if(bottomCell != null && ExtractUtil.getStringWithTrimmedText(bottomCell).length() > 0)
            count++;

        switch(count) {
            case 1:
                return "0,1,0,0,0";
            case 2:
                return "0,0,1,0,0";
            case 3:
                return "0,0,0,1,0";
            case 4:
                return "0,0,0,0,1";
            default:
                return "1,0,0,0,0";
        }
    }

    protected static String checkNeighbourType(Sheet sheet, Cell cell) {
        int rowNumber = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        CellType currentCellType = cell.getCellType();
        String res = "";

        Cell topCell = getCellIfExists(sheet, rowNumber-1, columnIndex);
        if(topCell != null && currentCellType.equals(topCell.getCellType()))
            res += "1,";
        else
            res += "0,";

        Cell bottomCell = getCellIfExists(sheet, rowNumber+1, columnIndex);
        if(bottomCell != null && currentCellType.equals(bottomCell.getCellType()))
            res += "1,";
        else
            res += "0,";

        Cell leftCell = getCellIfExists(sheet, rowNumber, columnIndex-1);
        if(leftCell != null && currentCellType.equals(leftCell.getCellType()))
            res += "1,";
        else
            res += "0,";

        Cell rightCell = getCellIfExists(sheet, rowNumber, columnIndex+1);
        if(rightCell != null && currentCellType.equals(rightCell.getCellType()))
            res += "1";
        else
            res += "0";

        return res;
    }

    protected static String checkNeighbourStyle(Sheet sheet, Cell cell) {
        int rowNumber = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        CellStyle currentCellStyle = cell.getCellStyle();
        String res = "";

        Cell topCell = getCellIfExists(sheet, rowNumber-1, columnIndex);
        if(topCell != null && currentCellStyle.equals(topCell.getCellStyle()))
            res += "1,";
        else
            res += "0,";

        Cell bottomCell = getCellIfExists(sheet, rowNumber+1, columnIndex);
        if(bottomCell != null && currentCellStyle.equals(bottomCell.getCellStyle()))
            res += "1,";
        else
            res += "0,";

        Cell leftCell = getCellIfExists(sheet, rowNumber, columnIndex-1);
        if(leftCell != null && currentCellStyle.equals(leftCell.getCellStyle()))
            res += "1,";
        else
            res += "0,";

        Cell rightCell = getCellIfExists(sheet, rowNumber, columnIndex+1);
        if(rightCell != null && currentCellStyle.equals(rightCell.getCellStyle()))
            res += "1";
        else
            res += "0";

        return res;
    }

    protected static String getNeighbourType(Sheet sheet, Cell cell) {
        int rowNumber = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        String res = "";

        Cell topCell = getCellIfExists(sheet, rowNumber-1, columnIndex);
        if(topCell != null)
            res += topCell.getCellType().getCode() + ",";
        else
            res += "-1,";

        Cell bottomCell = getCellIfExists(sheet, rowNumber+1, columnIndex);
        if(bottomCell != null)
            res += bottomCell.getCellType().getCode() + ",";
        else
            res += "-1,";

        Cell leftCell = getCellIfExists(sheet, rowNumber, columnIndex-1);
        if(leftCell != null)
            res += leftCell.getCellType().getCode() + ",";
        else
            res += "-1,";

        Cell rightCell = getCellIfExists(sheet, rowNumber, columnIndex+1);
        if(rightCell != null)
            res += rightCell.getCellType().getCode();
        else
            res += "-1";

        return res;
    }

}
