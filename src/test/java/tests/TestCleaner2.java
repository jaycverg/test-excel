package tests;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class TestCleaner2
{

    @Test
    public void test() throws Exception
    {
        String srcDirectory = "C:\\Users\\jvergara\\Documents\\Tests\\template-for-cleanup\\";
        String srcFilename = "Unaterra Shipping Lanes RFP data 02.08.13.xlsx";
//        String srcFilename = "HWL CUSTOMER 2 TEMPLATE - AIR FREIGHT.xls";
        File file = new File(srcDirectory + srcFilename);
        
        String destFilename = "d:\\test." + srcFilename.substring(srcFilename.lastIndexOf('.') + 1);

        Workbook workbook = WorkbookFactory.create(new FileInputStream(file));

        Sheet sheet = workbook.getSheet("RFP Worksheet");
//        Sheet sheet = workbook.getSheet("Response");
        for (int i=5; i<sheet.getLastRowNum(); ++i) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }

        sheet.shiftRows(6, sheet.getLastRowNum(), -1);
        

        OutputStream out = null;
        try {
            File destFile = new File(destFilename);
            workbook.write(out = new FileOutputStream(destFile));
            
            Runtime.getRuntime().exec(String.format("cmd /c \"%s\"", destFile.getAbsolutePath()));
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        finally {
            IOUtils.closeQuietly(out);
        }
    }
}
