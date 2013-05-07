package tests;

import org.junit.Test;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class JaroWinklerTest 
{
    @Test
    public void test()
    {
        System.out.println(JaroWinkler.compare("city of origin", "origin city"));
        System.out.println(JaroWinkler.compare("city of origin", "destination city"));
        System.out.println("---------");
        System.out.println(JaroWinkler.compare("port of origin", "origin port"));
        System.out.println(JaroWinkler.compare("port of origin", "destination port"));        
        System.out.println("---------");
        System.out.println(JaroWinkler.compare("country of origin", "origin country"));
        System.out.println(JaroWinkler.compare("country of origin", "destination country"));
        System.out.println("---------");
        System.out.println(JaroWinkler.compare("port of destination", "destination port"));
        System.out.println(JaroWinkler.compare("port of destination", "origin port"));
        System.out.println("---------");
        System.out.println(JaroWinkler.compare("city of destination", "destination city"));
        System.out.println(JaroWinkler.compare("city of destination", "origin city"));
        System.out.println("---------");
        System.out.println(JaroWinkler.compare("country of destination", "destination country"));
        System.out.println(JaroWinkler.compare("country of destination", "origin country"));
        System.out.println("---------");       
    }
}
