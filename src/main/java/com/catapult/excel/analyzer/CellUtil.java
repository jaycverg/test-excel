package com.catapult.excel.analyzer;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class CellUtil
{

    public static Object getCellValue(Cell cell)
    {
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC:
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_ERROR:
                return cell.getErrorCellValue();
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
        return (getColor(cell) != null);
    }

    public static Color getColor(Cell cell)
    {
        CellStyle style = cell.getCellStyle();
        return style.getFillForegroundColorColor();
    }

    public static boolean colorMatches(Cell cell1, Cell cell2)
    {
        if (cell1 == null || cell2 == null) {
            return false;
        }

        Color c1 = getColor(cell1);
        Color c2 = getColor(cell2);
        if (c1 instanceof XSSFColor || c2 instanceof XSSFColor) {
            XSSFColor xc1 = (XSSFColor) c1;
            XSSFColor xc2 = (XSSFColor) c2;

            return xc1.getARGBHex().equals(xc2.getARGBHex());
        }
        else if (c1 != null && c2 != null) {
            HSSFColor hc1 = (HSSFColor) c1;
            HSSFColor hc2 = (HSSFColor) c2;

            return hc1.getHexString().equals(hc2.getHexString());
        }

        return false;
    }

    public static boolean isTextBold(Cell cell)
    {
        CellStyle style = cell.getCellStyle();
        Font font = cell.getSheet().getWorkbook().getFontAt(style.getFontIndex());
        return (font.getBoldweight() == Font.BOLDWEIGHT_BOLD);
    }

    public static boolean hasBorder(Cell cell)
    {
        CellStyle style = cell.getCellStyle();
        short bt = style.getBorderTop();
        short br = style.getBorderRight();
        short bb = style.getBorderBottom();
        short bl = style.getBorderLeft();

        return  bt != CellStyle.BORDER_NONE ||
                br != CellStyle.BORDER_NONE ||
                bb != CellStyle.BORDER_NONE ||
                bl != CellStyle.BORDER_NONE;
    }
}
