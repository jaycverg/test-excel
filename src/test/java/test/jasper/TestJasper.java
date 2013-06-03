package test.jasper;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import net.sf.jasperreports.engine.JRDataSource;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jasperreports.engine.util.JRLoader;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class TestJasper 
{
    @Test
    public void test() throws Exception
    {
        Map m = new HashMap();
        m.put("field", "Sample Data");
        List list = new ArrayList();
        list.add(m);
        
        InputStream is = getClass().getResourceAsStream("Report1.jasper");
        JasperReport report = (JasperReport) JRLoader.loadObject(is);
        JRDataSource ds = new JRBeanCollectionDataSource(list);
        JasperPrint jp = JasperFillManager.fillReport(report, null, ds);
        JasperExportManager.exportReportToPdfFile(jp, "d:/test.pdf");

        Runtime.getRuntime().exec("explorer d:\\test.pdf");
    }

}
