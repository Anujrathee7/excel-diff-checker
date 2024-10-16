package edu.abhi.poi.excel;

import static org.junit.Assert.*;
import static org.mockito.Mockito.*;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.Future;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class ExcelDiffCheckerTest {

    @Mock
    private XSSFWorkbook mockWorkbook1;

    @Mock
    private XSSFWorkbook mockWorkbook2;

    @Mock
    private XSSFSheet mockSheet1;

    @Mock
    private XSSFSheet mockSheet2;

    @Before
    public void setUp() {
        MockitoAnnotations.openMocks(this);

        // Mocking the behavior of the workbooks and sheets
        when(mockWorkbook1.getSheet("Sheet1")).thenReturn(mockSheet1);
        when(mockWorkbook2.getSheet("Sheet1")).thenReturn(mockSheet2);
        when(mockWorkbook1.getSheet("SheetMissing")).thenReturn(null);
        when(mockWorkbook2.getSheet("SheetMissing")).thenReturn(null);
    }

    @Test
    public void testProcessSpecifiedSheets() {
        // Mock the specified sheets
        ExcelDiffChecker.specifiedSheets = "Sheet1,SheetMissing";
        List<SheetProcessorTask> tasks = ExcelDiffChecker.processSpecifiedSheets(mockWorkbook1, mockWorkbook2);

        // Verify the task count and behavior
        assertEquals(1, tasks.size());
    }

    @Test
    public void testProcessAllSheets() {
        when(mockWorkbook1.iterator()).thenReturn(Arrays.asList(mockSheet1).iterator());

        List<SheetProcessorTask> tasks = ExcelDiffChecker.processAllSheets(mockWorkbook1, mockWorkbook2);

        // There should be one task since both Sheet1 are present
        assertEquals(1, tasks.size());
    }

    @Test
    public void testExecuteAllTasks() throws Exception {
        SheetProcessorTask task = mock(SheetProcessorTask.class);
        List<SheetProcessorTask> tasks = Arrays.asList(task);

        Future<CallableValue> future = mock(Future.class);
        when(future.get()).thenReturn(new CallableValue(true));
        when(task.call()).thenReturn(new CallableValue(true));

        List<Future<CallableValue>> results = ExcelDiffChecker.executeAllTasks(tasks);

        // Check the result from future
        assertEquals(1, results.size());
    }

    @Test
    public void testProcessOptionsWithValidArgs() {
        String[] args = {"-b", "base.xlsx", "-t", "target.xlsx"};
        ExcelDiffChecker.processOptions(args);

        assertEquals("base.xlsx", ExcelDiffChecker.FILE_NAME1);
        assertEquals("target.xlsx", ExcelDiffChecker.FILE_NAME2);
    }

    @Test(expected = RuntimeException.class)
    public void testProcessOptionsWithMissingRequiredArgs() {
        String[] args = {"-b", "base.xlsx"};
        ExcelDiffChecker.processOptions(args);
    }
}
