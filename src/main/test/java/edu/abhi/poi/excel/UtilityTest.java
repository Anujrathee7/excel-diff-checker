package edu.abhi.poi.excel;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.IOException;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class UtilityTest {
    private XSSFCell cellEx;
	private Sheet sheetEx;
	private Workbook workbookEx;
	private CreationHelper crHeEx;

	// Set up example starting
	@BeforeEach
	void setUp() {
    		cellEx = mock(XSSFCell.class);
    		sheetEx = mock(Sheet.class);
    		workbookEx = mock(Workbook.class);
    		crHeEx = mock(CreationHelper.class);
    	when(cellEx.getSheet()).thenReturn(sheetEx);
    	when(sheetEx.getWorkbook()).thenReturn(workbookEx);
    	when(workbookEx.getCreationHelper()).thenReturn(crHeEx);
	}

	// Test to get a cell value of an excel file
 	@Test
	void testGetCellValue() throws Exception {
    	when(cellEx.getCellType()).thenReturn(CellType.STRING);
    	when(cellEx.getRichStringCellValue()).thenReturn(new XSSFRichTextString("example"));
     	String value = Utility.getCellValue(cellEx);
        	assertEquals("example", value);
	}
}
