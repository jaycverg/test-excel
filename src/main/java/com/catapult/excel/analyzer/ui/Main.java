package com.catapult.excel.analyzer.ui;

import com.catapult.excel.analyzer.ExcelDataHeader;
import com.catapult.excel.analyzer.ExcelSheetAnalyzer;
import java.io.InputStream;
import javax.swing.JDialog;
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
        new Main().testWithUI();
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
