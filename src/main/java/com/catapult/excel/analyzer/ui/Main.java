package com.catapult.excel.analyzer.ui;

import com.catapult.excel.analyzer.ExcelDataHeader;
import com.catapult.excel.analyzer.ExcelSheetAnalyzer;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import javax.swing.JFrame;
import javax.swing.UIManager;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.IOUtils;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class Main
{
    public static void main(String[] args)
    {
//        new Main().testWithUI();
        new Main().test();
    }

    private void testWithUI()
    {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        }
        catch(Exception e){}

        JFrame d = new JFrame("Test Excel Sheet Analyzer");
        d.setContentPane(new AnalyzerUI());
        d.pack();
        d.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        d.setVisible(true);
    }

    private void test()
    {
        InputStream is = null;
        try
        {
//            File f = new File("C:\\Users\\jvergara\\Documents\\Tests\\header-analyzer\\2013 RFP Matrix 02-19-13 with Hazmat.xlsx");
//            is = new FileInputStream(f);
//            Workbook wb = WorkbookFactory.create(is);
//            Sheet sheet = wb.getSheet("Expedited");

//            is = getClass().getResourceAsStream("test1.xlsx");
//            Workbook wb = WorkbookFactory.create(is);
//            Sheet sheet = wb.getSheetAt(0);

            is = getClass().getResourceAsStream("test1.xlsx");
            Workbook wb = WorkbookFactory.create(is);
            Sheet sheet = wb.getSheetAt(1);

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
