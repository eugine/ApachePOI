package com.eugenesokolov.apache.poi;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.*;

public class ApachePoiTest {

	@Test
	public void shouldCompile() {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet();
		
		sheet.createRow(0);
	}
}
