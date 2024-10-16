package edu.abhi.poi.excel;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.IOException;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

public class UtilityTest {
    private File temporaryFile;

    @Before
    public void initialize() throws IOException {
        temporaryFile = File.createTempFile("prefix-", "-suffix");
    }

    @After
    public void cleanup() {
        temporaryFile.delete();
    }

    @Test
    public void verifyDeletionOfNonExistingFile() {
        assertTrue(temporaryFile.delete());
        boolean isDeleted = Utility.deleteIfExists(temporaryFile);
        assertFalse(isDeleted);
    }

    @Test
    public void verifyDeletionOfExistingFile() {
        boolean isDeleted = Utility.deleteIfExists(temporaryFile);
        assertTrue(isDeleted);
    }
}
