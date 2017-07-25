package com.xlsxEditor;

import com.xlsxEditor.controller.myWorkbook;
import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Unit test for simple Driver.
 */
public class myWorkbookTest extends TestCase {
    XSSFWorkbook workBook;
    
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public myWorkbookTest(String testName) {
        super(testName);
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite() {
        return new TestSuite(myWorkbookTest.class);
    }

    public void setUp() throws Exception {
        super.setUp();
        workBook = new myWorkbook("C:/Users/crosenlund/Documents/xlsxEditor/src/test/resources/test.xlsx").getWorkBook();
        
    }
    
    /**
     * Test constructors
     */
    public void testMyWorkbook() {
        assertNotNull(workBook);
    }
}
