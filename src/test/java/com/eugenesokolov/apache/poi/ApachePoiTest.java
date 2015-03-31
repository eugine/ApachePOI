package com.eugenesokolov.apache.poi;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.dbunit.Assertion;
import org.dbunit.dataset.excel.XlsDataSet;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.TemporaryFolder;

public class ApachePoiTest {
    private static final int[][] EXPECTED_TEST_XLS_DATA = { { 1, 4, 7 }, { 2, 5, 8 }, { 3, 6, 9 } };
    private static final String[][] NEW_XLS_DATA = { { "A", "B", "C" }, { "1", "4", "7" }, { "2", "5", "8" }, { "3", "6", "9" } };

    @Rule
    public TemporaryFolder outputFolder = new TemporaryFolder();
    
    @Test
    public void loadXLSFile() throws FileNotFoundException, IOException {
        HSSFWorkbook workbook = loadWorkbook("/test.xls");
        HSSFSheet sheet = workbook.getSheetAt(0);

        assertSheet(EXPECTED_TEST_XLS_DATA, sheet);
    }

    @Test
    public void createXLSFile() throws Exception {
        HSSFWorkbook workbook = createWorkbookWithData(NEW_XLS_DATA);
        File outputFile = outputFolder.newFile("new.xls");
        saveWorkbook(workbook, outputFile);
        compareExcelFiles("/new_expected.xls", outputFile.getAbsolutePath());
    }

    private void saveWorkbook(HSSFWorkbook workbook, File outputFile) throws Exception {
        FileOutputStream out = new FileOutputStream(outputFile);
        workbook.write(out);
        out.close();
    }

    private HSSFWorkbook createWorkbookWithData(String[][] data) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sheet1");

        for (int i = 0; i < data.length; i++) {
            HSSFRow row = sheet.createRow(i);
            for (int j = 0; j < data[i].length; j++) {
                HSSFCell cell = row.createCell(j);
                cell.setCellValue(data[i][j]);
            }
        }
        return workbook;
    }

    private void compareExcelFiles(String expectedFile, String actualFile) throws Exception {
        File expected = getFileResource(expectedFile);
        File actual = new File(actualFile);
        Assertion.assertEquals(new XlsDataSet(expected), new XlsDataSet(actual));
    }

    private void assertSheet(int[][] expectedData, HSSFSheet actualSheet) {
        for (int i = 0; i < expectedData.length; i++) {
            HSSFRow row = actualSheet.getRow(i);
            for (int j = 0; j < expectedData[i].length; j++) {
                assertEquals(expectedData[i][j], row.getCell(j).getNumericCellValue(), 0.001);
            }
        }
    }

    private HSSFWorkbook loadWorkbook(String name) throws IOException {
        File file = getFileResource(name);
        return new HSSFWorkbook(new FileInputStream(file));
    }

    private File getFileResource(String name) {
        URL url = this.getClass().getResource(name);
        File file = new File(url.getPath());
        return file;
    }

}
