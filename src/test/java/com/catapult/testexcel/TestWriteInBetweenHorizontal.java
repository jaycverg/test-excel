package com.catapult.testexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class TestWriteInBetweenHorizontal 
{
    @Test
    public void test() throws Exception
    {
        File source = new File("test-resources/template2.xlsx");
        File dest = new File("test-resources/template2_modified.xlsx");
                
        Workbook wb = WorkbookFactory.create(source);
        
        Sheet sheet = wb.getSheetAt(0);
                        
        int startRow = 1;
        int endRow = 10;
        int startCol = 1;
        int nextHeaderCol = 3;
        int colsAdded = 0;
        int groupGap = 2;
        
        //lets keep the template cells' style
        Map<Integer, CellStyle> cellStyles = new HashMap();
        
        //keep the next headers
        Map<Integer, CellClone> nextHeaders = new HashMap();
        
        for (int rowIdx = startRow; rowIdx <= endRow; ++rowIdx)
        {
            Row row = sheet.getRow(rowIdx);
            CellStyle cs = row.getCell(startCol).getCellStyle();
            cellStyles.put(rowIdx, cs);
            
            //get the next header colIdx
            Cell cell = row.getCell(nextHeaderCol);
            nextHeaders.put(rowIdx, new CellClone(cell));
        }
        
        //load the data for the first data group
        for (int colIdx = startCol, colCount = 0; colCount < 5; ++colIdx, ++colCount)
        {
            sheet.setColumnWidth(colIdx, 5000);
            for (int rowIdx = startRow; rowIdx <= endRow; ++rowIdx) {
                Cell cell = sheet.getRow(rowIdx).createCell(colIdx);
                cell.setCellStyle(cellStyles.get(rowIdx));
                cell.setCellValue(String.format("Data: %s, %s", rowIdx, colIdx));
            }
            colsAdded++;
        }
        
        //re-draw the next headers
        sheet.setColumnWidth(startCol + colsAdded + groupGap, 5000);
        for (int rowIdx = startRow; rowIdx <= endRow; ++rowIdx) {
            Row row = sheet.getRow(rowIdx);
            Cell newCell = row.createCell(startCol + colsAdded + groupGap);
            CellClone clone = nextHeaders.get(rowIdx);
            clone.copyTo(newCell);
        }
        colsAdded++;
        
        //load the data for the next data group        
        for (int colIdx = startCol + colsAdded + groupGap, colCount = 0; colCount < 5; ++colIdx, ++colCount)
        {
            sheet.setColumnWidth(colIdx, 5000);
            for (int rowIdx = startRow; rowIdx <= endRow; ++rowIdx) {
                Cell cell = sheet.getRow(rowIdx).createCell(colIdx);
                cell.setCellStyle(cellStyles.get(rowIdx));
                cell.setCellValue(String.format("Data: %s, %s", rowIdx, colIdx));
            }
            colsAdded++;
        }
        
        wb.write(new FileOutputStream(dest));
        
        //display the output
        Runtime.getRuntime().exec("cmd /c " + dest.getAbsolutePath());
    }
    
    private class CellClone 
    {
        private Object value;
        private int cellType;
        private CellStyle cellStyle;
        private Comment comment;

        CellClone(Cell cell)
        {
            cellType = cell.getCellType();
            cellStyle = cell.getCellStyle();
            comment = cell.getCellComment();
            
            switch (cellType) {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    value = cell.getNumericCellValue();
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    value = null;
                    break;
                case Cell.CELL_TYPE_ERROR:
                    value = cell.getErrorCellValue();
                    break;
            }
        }
        
        void copyTo(Cell cell)
        {
            cell.setCellStyle(cellStyle);
            cell.setCellComment(comment);
            
            switch (cellType) {
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
    }
}
