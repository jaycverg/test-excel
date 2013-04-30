package com.catapult.testexcel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class TestWriteInBetween 
{
    @Test
    public void test() throws Exception
    {
        File source = new File("test-resources/template.xlsx");
        File dest = new File("test-resources/template_modified.xlsx");
                
        Workbook wb = WorkbookFactory.create(source);
        
        Sheet sheet = wb.getSheetAt(0);
                
        //get the template row columns
        Row row = sheet.getRow(2);
        List<CellStyle> cellStyles = new ArrayList();
        for (int c = 0; c < 4; ++c) {
            cellStyles.add(row.getCell(c).getCellStyle());
        }
        
        for (int r = 0; r < 5; ++r) {
            sheet.shiftRows(r+2, sheet.getLastRowNum(), 1);
            row = sheet.createRow(r+2);
            for (int c = 0; c < 4; ++c) {
                Cell cell = row.createCell(c);
                cell.setCellStyle(cellStyles.get(c));
                cell.setCellValue("Column " + c);
            }
        }
        
        wb.write(new FileOutputStream(dest));
        
        //display the output
        Runtime.getRuntime().exec("cmd /c " + dest.getAbsolutePath());
    }
}
