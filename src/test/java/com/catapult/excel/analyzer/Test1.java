package com.catapult.excel.analyzer;

import java.text.MessageFormat;
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
        MessageFormat mf = new MessageFormat("Hello {1} and {0}");
        System.out.println(mf.format(new Object[]{"jay", "alvin"}));
    }
}
