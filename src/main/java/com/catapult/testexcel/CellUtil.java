package com.catapult.testexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class CellUtil 
{
    public static Object getCellValue(Cell cell)
    {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN: return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC: return cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING:  return cell.getStringCellValue();
            case Cell.CELL_TYPE_ERROR:   return cell.getErrorCellValue();
        }
        return null;
    }

    public static String getCellValueAsString(Cell cell)
    {
        Object value = getCellValue(cell);
        return (value == null) ? "" : value.toString().trim();
    }
    
    public static void setCellValue(Cell cell, Object value)
    {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue((Boolean) value);
                break;
            case Cell.CELL_TYPE_NUMERIC:
                cell.setCellValue((Double) value);
                break;
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue((String) value);
                break;
            case Cell.CELL_TYPE_BLANK:
                cell.setCellValue("");
                break;
            case Cell.CELL_TYPE_ERROR:
                cell.setCellErrorValue((Byte) value);
                break;
        }
    }

    public static boolean hasBackground(Cell cell)
    {
        CellStyle style = cell.getCellStyle();
        return (style.getFillForegroundColorColor() != null);
    }

    public static boolean isTextBold(Cell cell)
    {
        CellStyle style = cell.getCellStyle();
        Font font = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());
        return (font.getBoldweight() == Font.BOLDWEIGHT_BOLD);
    }
}
