package tests;

import java.io.File;
import java.io.FilenameFilter;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class Test1 
{
    private FilenameFilter excelFileFilter = new FilenameFilter() {

        public boolean accept(File dir, String name)
        {
            return name.endsWith(".xls")
                    || name.endsWith(".xlsx")
                    || name.endsWith(".xlsm");
        }
    };

    @Test
    public void test()
    {
        File d = new File("C:/Users/jvergara/Documents/Tests/template-cleaned");
        for (File f : d.listFiles(excelFileFilter)) {
            System.out.println(f.getName());
        }
    }
}
