/*
 * Copyright (c) 2019  Dubai Police and/or its affiliates. All rights reserved.
 * DUBAI POLICE PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 *
 */

package export.excel.common.customImplementation;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import export.excel.common.model.ExcelCustomAdd;

/**
 * 
 * The {@code CustomParamGoodConductRpt} implements custom param implementation
 * for excel. It extends {@code ExcelCustomAdd} that provides common methods to
 * work on excel.
 * 
 * @author Ruba Mushtaq
 * @see export.excel.common.model.ExcelCustomAdd
 * @since JDK1.5
 */

public class CustomParamImplRpt extends ExcelCustomAdd {

	public CustomParamImplRpt(Map<String, Object> customHeader, Map<String, Object> customFooter,
			Map<String, Object> customParams) {
		super(customHeader, customFooter, customParams);
	}

	/**
	 * Reads value from {@code customParms} , set it on excel by calling parent method.
	 * To format rows and column , it calls standard library.
	 *  
	 */
	public void customParams() {
		HSSFRow row = this.addRow();

		int columnSize = (int) customParams.get("columnSize");

		HSSFCellStyle rowCellStyle = this.getExcelSheetEdit().getWorkbook().createCellStyle();
		byte[] rgb = new byte[3];
		rgb[0] = (byte) 217; // red
		rgb[1] = (byte) 225; // blue
		rgb[2] = (byte) 242; // green
		HSSFColor lightBlue = setColor(this.getExcelSheetEdit().getWorkbook(), HSSFColor.BLUE.index, rgb[0], rgb[1],
				rgb[2]);
		HSSFFont rowHeadingFont = this.getExcelSheetEdit().getWorkbook().createFont();
		rowHeadingFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		rowHeadingFont.setCharSet(HSSFFont.ANSI_CHARSET);
		rowHeadingFont.setFontHeightInPoints((short) 11);
		rowHeadingFont.setColor(IndexedColors.BLACK.getIndex());
		rowHeadingFont.setFontName("Arial");
		rowHeadingFont.setUnderline(HSSFFont.U_SINGLE);
		rowCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		rowCellStyle.setFillForegroundColor(lightBlue.getIndex());
		rowCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND); // header
		rowCellStyle.setFont(rowHeadingFont);
		String formatVal = (String) customParams.get("param1") + customParams.get("param2")
				+ customParams.get("natType") + customParams.get("statusName");
		HSSFCell cell1 = this.addCell(row, 0, formatVal);
		cell1.setCellStyle(rowCellStyle);
		this.getExcelSheetEdit().getSheet()
				.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), 0, columnSize * 2 - 1));

	}

	public static HSSFColor setColor(HSSFWorkbook workbook, short color, byte r, byte g, byte b) {
		HSSFPalette palette = workbook.getCustomPalette();
		HSSFColor hssfColor = null;
		try {
			hssfColor = palette.findColor(r, g, b);
			if (hssfColor == null) {
				palette.setColorAtIndex(color, r, g, b);
				hssfColor = palette.getColor(color);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return hssfColor;
	}

	private static final long serialVersionUID = 1L;

}
