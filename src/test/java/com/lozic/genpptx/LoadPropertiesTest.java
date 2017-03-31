package com.lozic.genpptx;

import org.junit.Test;
import org.junit.Before;
import org.junit.After;

import java.util.List;

/**
 * LoadProperties Tester.
 *
 * @author <Authors name>
 * @version 1.0
 * @since <pre>十一月 22, 2016</pre>
 */
public class LoadPropertiesTest {

    @Before
    public void before() throws Exception {
    }

    @After
    public void after() throws Exception {
    }

    /**
     * Method: getPPTBean()
     */
    @Test
    public void testGetPPTBean() throws Exception {
        LoadProperties lp = new LoadProperties();
        //PPTBean pptBean = lp.getPPTBean();
        //pptBean.getSlideEntitiesList();
    }

    @Test
    public void testCreatePPTBean(){


    }

    @Test
    public void testGetAllLines() {
        LoadProperties lp = new LoadProperties();
        List<String> ll = lp.getAllLines("/Users/wangqingwu/Projects/gen-pptx/ppcloud.properties");

        ll.stream().forEach(line -> System.out.println(line));
    }

} 
