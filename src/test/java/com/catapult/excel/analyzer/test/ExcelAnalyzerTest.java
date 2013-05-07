package com.catapult.excel.analyzer.test;

import com.catapult.excel.analyzer.ExcelDataHeader;
import com.catapult.excel.analyzer.ExcelSheetAnalyzer;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class ExcelAnalyzerTest
{
    @Test
    public void test()
    {
        InputStream is = null;
        try
        {
//            is = getClass().getResourceAsStream("format4.xlsx");
//            Workbook wb = WorkbookFactory.create(is);
//            Sheet sheet = wb.getSheetAt(wb.getSheetIndex("Rates"));

            is = getClass().getResourceAsStream("format2.xlsx");
            Workbook wb = WorkbookFactory.create(is);
            Sheet sheet = wb.getSheetAt(0);

            System.out.println("-------------------------------------");
            System.out.println("Sheet Name: " + sheet.getSheetName());
            System.out.println("-------------------------------------");

            ExcelSheetAnalyzer esa = new ExcelSheetAnalyzer(sheet);
            esa.analyze();
            for (ExcelDataHeader header : esa.getHeaders())
            {
                System.out.println(header);
            }
        }
        catch(Exception e) {
            e.printStackTrace();
        }
        finally {
            IOUtils.closeQuietly(is);
        }
    }
}
