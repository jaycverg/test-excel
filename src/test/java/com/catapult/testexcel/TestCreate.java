package com.catapult.testexcel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class TestCreate 
{
    @Test
    public void test() throws Exception
    {
        String destFilename = "test-resources/custom_template.xlsx";
        String logoFilename = "test-resources/catapult.png";
        
        Workbook wb = new XSSFWorkbook();
        
        Sheet sheet = wb.createSheet("Sheet 1");
                
        //header company title
        Row row = sheet.createRow(0);
        row.setHeightInPoints(50f);
        
        Cell cell = row.createCell(0);
        cell.setCellValue("Catapult International, LLC.");
        
        CellStyle css = wb.createCellStyle();
        css.setFillForegroundColor(IndexedColors.DARK_TEAL.index);
        css.setFillPattern(CellStyle.SOLID_FOREGROUND);
        css.setAlignment(CellStyle.ALIGN_CENTER);
        css.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font font = wb.createFont();
        font.setColor(IndexedColors.WHITE.index);
        font.setFontHeightInPoints((short) 18);
        css.setFont(font);
        
        cell.setCellStyle(css);
        
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
        
        //add the company logo
        byte[] bytes = loadBytesFromFile(new FileInputStream(logoFilename));
        int picIndex = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        
        CreationHelper creationHelper = wb.getCreationHelper();
        Drawing drawing = sheet.createDrawingPatriarch();
        
        ClientAnchor anchor = creationHelper.createClientAnchor();
        anchor.setCol1(0);
        anchor.setRow1(0);
        anchor.setAnchorType(ClientAnchor.DONT_MOVE_AND_RESIZE);
        
        Picture pic = drawing.createPicture(anchor, picIndex);
        pic.resize();
        
        //specify column widths
        for (int c = 0; c < 4; ++c) {
            sheet.setColumnWidth(c, 6500);
        }
        
        
        //default header Cell Style
        CellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREEN.index);
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
        headerStyle.setBorderTop(CellStyle.BORDER_THIN);
        headerStyle.setBorderRight(CellStyle.BORDER_THIN);
        headerStyle.setBorderBottom(CellStyle.BORDER_THIN);
        headerStyle.setBorderLeft(CellStyle.BORDER_THIN);
        font = wb.createFont();
        font.setColor(IndexedColors.WHITE.index);
        headerStyle.setFont(font);
        
        //data item Cell Style
        CellStyle itemStyle = wb.createCellStyle();
        itemStyle.setBorderTop(CellStyle.BORDER_THIN);
        itemStyle.setBorderRight(CellStyle.BORDER_THIN);
        itemStyle.setBorderBottom(CellStyle.BORDER_THIN);
        itemStyle.setBorderLeft(CellStyle.BORDER_THIN);
        
        //lets create two data groups
        int currentRow = 1;
        for (int group=0; group < 4; ++group) 
        {            
            //create the row headers
            row = sheet.createRow(currentRow);
            int hGroupCount = 0;
            for (int c = 0; c < 4; ++c) {
                cell = row.createCell(c);
                cell.setCellStyle(headerStyle);
                if (c % 2 == 0) {
                    cell.setCellValue("Header Group " + ++hGroupCount);
                    int cellStart = (hGroupCount-1)*2;
                    sheet.addMergedRegion(new CellRangeAddress(currentRow, currentRow, cellStart, cellStart+1));
                }
            }
            
            row = sheet.createRow(++currentRow);
            for (int c = 0; c < 4; ++c) {
                cell = row.createCell(c);
                cell.setCellValue("Header " + (c+1));
                cell.setCellStyle(headerStyle);
            }

            //insert dummy data
            for (int r = 0; r < 10; ++r) {
                row = sheet.createRow(++currentRow);
                for (int c = 0; c < 4; ++c) {
                    cell = row.createCell(c);
                    cell.setCellValue(String.format("Data: %s, %s", currentRow, c+1));
                    cell.setCellStyle(itemStyle);
                }
            }
            
            currentRow += 3;
        }
        
        wb.write(new FileOutputStream(destFilename));
        
        //display the output
        Runtime.getRuntime().exec("cmd /c " + new File(destFilename).getAbsolutePath());
    }
    
    private byte[] loadBytesFromFile(InputStream is)
    {
        try {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            int i = -1;
            while ((i = is.read()) != -1) {
                bos.write(i);
            }
            return bos.toByteArray();
        }
        catch(Exception e) {
            throw new RuntimeException(e);
        }
        finally {
            try { is.close(); }catch(Exception e){}
        }
    }
    
}
