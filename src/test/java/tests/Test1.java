package tests;

import java.io.File;
import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class Test1 
{
    @Test
    public void test()
    {
        File d = new File("C:/Users/jvergara/Documents/Tests/header-analyzer");
        for (File f : d.listFiles()) {
            System.out.println(f.getName());
        }
    }
}
