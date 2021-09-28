package org.commandline.exceltojson;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExcelToJSONTest {

    private File testFile;

    @BeforeEach
    public void setUp() {
        ClassLoader classLoader = this.getClass().getClassLoader();
        testFile = new File(classLoader.getResource("FinancialSample.xlsx").getFile());
    }

    @Test
    public void testSetupIsFunctioning() {
        assertNotNull(testFile);
    }

    //VERY simple smoke test
    @Test
    public void testBasicConversion() {
        ExcelToJSON unit = new ExcelToJSON(testFile);
        String jsonOutput = unit.toJson();
        assertTrue(jsonOutput.contains("\"Row 1\""));
        assertTrue(jsonOutput.contains("United States of America"));
        assertTrue(jsonOutput.contains("\"Row 700\""));
    }


}
