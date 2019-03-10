/*
 * Copyright (c) 2019
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

package export.excel.common.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import org.apache.commons.io.IOUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;

import export.excel.common.constants.ExcelConstant;
import export.excel.common.dto.ExcelDTO;
import export.excel.common.dto.ExcelPrintDTO;
import export.excel.common.model.ExcelCustomAdd;
import export.excel.common.model.ExcelSheetEdit;

/**
 * 
 * The {@code ExcelExportManager} is an utility class that generates excel and
 * return {@code ExcelPrintDTO} that contains excel file information.Provides
 * {@code exportExcel} methods that generates excel default and
 * custom implementation as well providing data index to be arranged in Excel.
 * 
 * @author Ruba Mushtaq
 * @see org.apache.poi.hssf.usermodel.HSSFWorkbook
 * @see org.apache.poi.hssf.usermodel.HSSFSheet
 * @see export.excel.common.model.ExcelSheetEdit
 * @see export.excel.common.dto.ExcelPrintDTO
 * @since JDK1.5
 */
public class ExcelExportManager {

	/** maintains row count in excel */
	public volatile static Integer rowNum;
	
	/** The value is used for excel workbook storage */
	private static HSSFWorkbook workbook;
	
	/** The value is used for excel work sheet storage in {@code workbook} */
	private static HSSFSheet sheet;
	
	/** The value is used for temporary excel sheet name*/
	private final static String EXCEL_FILE = "excel-temp.xls";
	
	/** used to format excel {@code ExcelSheetEdit}. Contains all necessary
	 * functionality to be worked on Excel*/
	private static ExcelSheetEdit excelSheetEdit;
	
	/** used to fix {@code HSSFSheet} style and format*/
	private static int rowListSize;
	
	/** used to set width of columns into a generated excel*/
	private final static int COLUMN_WIDTH = 6000;
	
	
	/**
	 * Allocate all requisite resources
	 * 
	 */
	private static void defualtRequiredSetup() {
		workbook = new HSSFWorkbook();
		sheet = workbook.createSheet(ExcelConstant.FIRST_SHEET.getValue());
		rowNum = -1;
		//sheet.setRightToLeft(true);
		excelSheetEdit = new ExcelSheetEdit(workbook, sheet, rowNum);
	}

	/**
	 * Return {@code ExcelPrintDTO} that contains generated excel information.
	 * Generates Excel with default implementation
	 * 
	 * @param defaultMap             contains header and footer information
	 * @param rowHeadingLS           Titles for columns to be displayed in Excel
	 * @param reportDTOLS            value for rows data to be shown in excel
	 * @param dataIndex              custom column indexes to be arranged in excel
	 * 								 if not provided index, the rows data will be arranged as
	 * 								 defined in {@code reportDTOLS} objects
	 * 
	 * @return                       {@code ExcelPrintDTO}
	 * @throws Exception
	 */
	public static ExcelPrintDTO exportExcel(EnumMap<ExcelConstant, Object> defaultMap, List<String> rowHeadingLS,
			List<ExcelDTO> reportDTOLS, int... dataIndex) throws Exception {

		try {
			ExcelExportManager.defualtRequiredSetup();
			rowListSize = rowHeadingLS.size();
			ExcelExportManager.perusal_default_map(defaultMap, true, false);
			createColumnTitle(workbook, sheet, rowHeadingLS);
			setDataIntoCell(reportDTOLS, dataIndex);
			ExcelExportManager.perusal_default_map(defaultMap, false, true);
			return createExcelPrintDTO();

		} catch (Exception e) {
			throw e;
		}

	}
	
	/**
	 * Return {@code ExcelPrintDTO} that contains generated excel information.
	 * Generates Excel with custom implementation
	 * 
	 * @param defaultMap             contains header and footer information
	 * @param rowHeadingLS           Titles for columns to be displayed in Excel
	 * @param reportDTOLS            value for rows data to be shown in excel
	 * @param dataIndex              custom column indexes to be arranged in excel
	 * 								 if not provided index, the rows data will be arranged as
	 * 								 defined in {@code reportDTOLS} objects
	 * 
	 * @return                       {@code ExcelPrintDTO}
	 * @throws Exception
	 */
	public static ExcelPrintDTO exportExcel(EnumMap<ExcelConstant, Object> defaultMap, List<String> rowHeadingLS,
			ExcelCustomAdd excelCustomAdd, List<ExcelDTO> reportDTOLS, int... dataIndex) throws Exception {
		try {

			ExcelExportManager.defualtRequiredSetup();
			excelCustomAdd.setExcelSheetEdit(excelSheetEdit);
			rowListSize = rowHeadingLS.size();
			ExcelExportManager.perusal_default_map(defaultMap, true, false);
			excelCustomAdd.customHeaderImpl();
			excelCustomAdd.customParams();
			createColumnTitle(workbook, sheet, rowHeadingLS);
			setDataIntoCell(reportDTOLS, dataIndex);
			excelCustomAdd.customFooterImpl();
			ExcelExportManager.perusal_default_map(defaultMap, false, true);
			return createExcelPrintDTO();
		} catch (Exception e) {
			throw e;
		}

	}

	/**
	 *  Scan {@code defaultMap} to write header and footer in excel
	 *  
	 * @param defaultMap      contains header and footer key values
	 * @param header          if true write headers into the sheet
	 * @param footer          if true write footers into the sheet
	 */
	private static void perusal_default_map(EnumMap<ExcelConstant, Object> defaultMap, boolean header, boolean footer) {
		if (defaultMap != null && defaultMap.size() > 0) {
			if (header) {
				defaultHeaders(defaultMap);
			} else if (footer) {
				defaultFooters(defaultMap);
			}
		}
	}
	
	/**
	 * Create and format Title row for data in excel 
	 * 
	 * @param workbook      used to create style for titles
	 * @param sheet         set width for column in the sheet
	 * @param columns       set column into cells created on {@code workbook} {@code sheet}
	 */
	private static void createColumnTitle(HSSFWorkbook workbook, HSSFSheet sheet, List<String> columns) {
		HSSFRow row = excelSheetEdit.addRow();
		mergeRegion(row.getRowNum(), row.getRowNum(), 0, (rowListSize * 2) - 1);
		HSSFRow colTitleRow = excelSheetEdit.addRow();

		HSSFCellStyle colTitleStyle = workbook.createCellStyle();
		HSSFCell lastCellWith = excelSheetEdit.addCell(colTitleRow, columns.size() * 2);
		sheet.setColumnWidth(lastCellWith.getColumnIndex(), 20000);
		setBorders(colTitleStyle);
		HSSFFont colTitleFont = getBoldFont(12, IndexedColors.BLACK.getIndex(), "Tahoma", HSSFFont.ANSI_CHARSET);
		colTitleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		colTitleStyle.setFont(colTitleFont);
		setForeground(colTitleStyle, IndexedColors.YELLOW.getIndex(), HSSFCellStyle.SOLID_FOREGROUND);
		int cellndex = -1;
		for (int i = 0; i < columns.size(); i++) {
			HSSFCell titleCell = excelSheetEdit.addCell(colTitleRow, ++cellndex, columns.get(i));
			HSSFCell mergeWith = excelSheetEdit.addCell(colTitleRow, ++cellndex);
			mergeRegion(colTitleRow.getRowNum(), colTitleRow.getRowNum(), titleCell.getColumnIndex(), mergeWith.getColumnIndex());
			titleCell.setCellStyle(colTitleStyle);
			mergeWith.setCellStyle(colTitleStyle);
			sheet.setColumnWidth(titleCell.getColumnIndex(), COLUMN_WIDTH);
		}

	}
	
	/**
	 * Generate data for rows
	 * 
	 * @param reportDTOLs           row data values to be placed in excel
	 * @param dataIndex             set {@code reportDTOLs} on index as defined in {@code dataIndex}
	 * @throws Exception
	 */
	private static void setDataIntoCell(List<ExcelDTO> reportDTOLs, int... dataIndex) throws Exception {

		if (reportDTOLs != null && reportDTOLs.size() > 0) {
			int[] dataPositions = do_up_index(reportDTOLs.get(0).getExcelDataLS().size(), dataIndex);
			for (ExcelDTO d : reportDTOLs) {
				HSSFRow row = excelSheetEdit.addRow();
				setRowData(d, row, dataPositions);

			}
		}
		HSSFRow rowMerge1 = excelSheetEdit.addRow();
		HSSFRow rowMerge2 = excelSheetEdit.addRow();
		mergeRegion(rowMerge1.getRowNum(), rowMerge2.getRowNum(),0, (rowListSize * 2) - 1);
		

	}
	
	/**
	 * Place data into rows and format it
	 * @param excelDTO              values to be placed in row
	 * @param row                   designated row to be set with {@code excelDTO}
	 * @param dataIndexes           set {@code excelDTO} on index as defined in {@code dataIndex}
	 * @throws Exception    
	 */
	private static void setRowData(ExcelDTO excelDTO, HSSFRow row, int[] dataIndex) throws Exception {

		HSSFCellStyle dataCellStyle = workbook.createCellStyle();
		setBorders(dataCellStyle);
		dataCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		HSSFFont dataFont = getFont(10, IndexedColors.BLACK.getIndex(), "Tahoma", HSSFFont.ANSI_CHARSET);
		dataCellStyle.setFont(dataFont);
		List<Object> cellValue = excelDTO.getExcelDataLS();
		int getValueFrom = 1;
		int count = dataIndex.length;
		int cellndex = -1;
		for (int listCounter = 1; listCounter < count; listCounter++) {
			int getValue = dataIndex[getValueFrom];
			HSSFCell dataCell = excelSheetEdit.addCell(row, ++cellndex, "" + cellValue.get(getValue));
			HSSFCell mergeWith = excelSheetEdit.addCell(row, ++cellndex);
			mergeRegion(row.getRowNum(), row.getRowNum(), dataCell.getColumnIndex(), mergeWith.getColumnIndex());
			dataCell.setCellStyle(dataCellStyle);
			mergeWith.setCellStyle(dataCellStyle);
			getValueFrom++;
		}

	}
	
	/**
	 *  Writes default header into excel
	 *  @param headers           contains headers key values
	 */
	public static void defaultHeaders(EnumMap<ExcelConstant, Object> headers) {
		if (ifContains(headers, true, false)) {
			createReportTiltle(headers);
			excelSheetEdit.setRowCount(rowNum);
			HSSFRow headerRow1 = excelSheetEdit.addRow();
			mergeRegion(headerRow1.getRowNum(), headerRow1.getRowNum(), 0, (rowListSize * 2) - 1);
			++rowNum;
			
			HSSFRow headerRow2 = excelSheetEdit.addRow();
			mergeRegion(headerRow2.getRowNum(), headerRow2.getRowNum(), 0, (rowListSize * 2) - 1);
			excelSheetEdit.setRowCount(rowNum);

			HSSFCellStyle headerCellStyle = workbook.createCellStyle();
			HSSFFont headerFont = getUnderLineBoldFont(14, IndexedColors.BLACK.getIndex(), "Arial",
					HSSFFont.ANSI_CHARSET, HSSFFont.U_SINGLE);
			headerCellStyle.setFont(headerFont);
			headerCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

			String fromDateTitle = ExcelConstant.FROM_DATE.getValue();
			String toDateTitle = ExcelConstant.TO_DATE.getValue();

			String fromDateVal = (String) headers.get(ExcelConstant.FROM_DATE);
			String toDateVal = (String) headers.get(ExcelConstant.TO_DATE);
			String formattedValue = fromDateTitle + ": " + fromDateVal + " " + toDateTitle + ": " + toDateVal;
			HSSFRow defaultHeaderRow = excelSheetEdit.addRow();
			HSSFCell headerCell = excelSheetEdit.addCell(headerRow2, 0, formattedValue);
			mergeRegion(defaultHeaderRow.getRowNum(), defaultHeaderRow.getRowNum(), headerCell.getColumnIndex(),
					(rowListSize * 2) - 1);
			headerCell.setCellStyle(headerCellStyle);

		}

	}
	
	/**
	 * Returns the after fixed dataIndex
	 * 
	 * @param noOfFields           no of column titles in excel
	 * @param dataIndex            no of fields in excel
	 * <blockquote><pre>
	 * 	correct result will be generated when noOfFields = dataIndex
	 *  if(noOfFields !=dataIndex){
	 *  	Unpredictable result may generate
	 *  }
	 *  </pre></blockquote>
	 * 
	 * @return <code>int[]</code>
	 * @throws Exception 
	 */
	private static int[] do_up_index(int noOfFields, int... dataIndex) throws Exception {
		int[] cellIndexes = dataIndex;
		
		int cellValueLength = noOfFields;
		int dataIndexLength = dataIndex != null ? dataIndex.length : 0;
		int startCount = 1;                     // o index is always escaped

		int dataIndexVal = 0;
		if (dataIndex == null || dataIndex.length == 0 || dataIndexLength > cellValueLength-1 || dataIndexLength < cellValueLength-1 ) {                    //reset to default index if any true
			cellIndexes = new int[cellValueLength];
			for (; startCount < cellValueLength; startCount++) {
				cellIndexes[startCount] = startCount;
			}

		}
		else  {
			cellIndexes = new int[noOfFields];
			int i = 0;
			for (; i < noOfFields - 1; i++, startCount++) {

				dataIndexVal = dataIndex[i];
				if (dataIndexVal == 1) {
					cellIndexes[startCount] = dataIndexVal;
				}
				if (dataIndexVal >cellValueLength - 1)
					throw new Exception(new ArrayIndexOutOfBoundsException(ExcelFilePropertyManager.getValue("INCORRECT.INDEX")));
				cellIndexes[startCount] = dataIndexVal;

			}
		}
			return cellIndexes;

	}
	
	/**
	 *  Writes default footers into excel
	 *  @param footers           contains footer key values
	 */
	private static void defaultFooters(EnumMap<ExcelConstant, Object> footers) {
		if (ifContains(footers, false, true)) {
			HSSFCellStyle footerCellStyle = workbook.createCellStyle();
			byte[] rgb = new byte[3];
			rgb[0] = (byte) 208; // red
			rgb[1] = (byte) 206; // green
			rgb[2] = (byte) 206; // blue

			HSSFColor lightTan = setColor(workbook, HSSFColor.GREY_25_PERCENT.index, rgb[0], rgb[1], rgb[2]);
			setForeground(footerCellStyle, lightTan.getIndex(), HSSFCellStyle.SOLID_FOREGROUND);

			HSSFCellStyle noBoldCell = workbook.createCellStyle();
			setForeground(noBoldCell, lightTan.getIndex(), HSSFCellStyle.SOLID_FOREGROUND);

			HSSFFont footerFont = getBoldFont(11, IndexedColors.BLACK.getIndex(), "Arial", HSSFFont.ANSI_CHARSET);
			footerCellStyle.setFont(footerFont);
			List<ExcelConstant> footerKeys = new ArrayList<>();
			footerKeys.add(ExcelConstant.REPORTED_BY);
			footerKeys.add(ExcelConstant.REPORTED_ON);
			int cellCounter = -1;
			int mapCounter = 0;
			for (; mapCounter < 2;) {
				cellCounter = -1;
				if (footers.containsKey(footerKeys.get(mapCounter))) {
					HSSFRow footerRow1 = excelSheetEdit.addRow();
					String title = (footerKeys.get(mapCounter)).getValue();
					Object value = footers.get(footerKeys.get(mapCounter));
					String formattedValue = title + " : " + value;

					HSSFCell footerCell1 = excelSheetEdit.addCell(footerRow1, ++cellCounter, formattedValue);
					footerCell1.setCellStyle(footerCellStyle);
					cellCounter = ++cellCounter + 2;
					mergeRegion(footerRow1.getRowNum(), footerRow1.getRowNum(), footerCell1.getColumnIndex(), cellCounter);
					HSSFCell footerCell2 = excelSheetEdit.addBlankCell(footerRow1, ++cellCounter);
					footerCell2.setCellStyle(noBoldCell);
					mergeRegion(footerRow1.getRowNum(), footerRow1.getRowNum(), footerCell2.getColumnIndex(),
							(rowListSize * 2) - 1);
				}
				mapCounter++;
			}
		}

	}
	
	/**
	 * Returns generated excel file after writing data into excel.
	 * 
	 * @return <code>File</code>
	 * @throws Exception
	 */
	private static File writeInFile() throws Exception {
		File file = null;
		FileOutputStream out = null;
		try {
			file = new File(".", ExcelExportManager.EXCEL_FILE);
			if (file.exists()) {
				file.delete();
				file = new File(".", ExcelExportManager.EXCEL_FILE);
			}

			out = new FileOutputStream(file);
			workbook.write(out);

		} catch (Exception e) {

			/*
			 * if (file != null && file.exists()) { file.delete(); }
			 */

			throw e;
		} finally {
			if (out != null) {
				out.close();
			}

		}

		return file;
	}

	/**
	 * Return {@code ExcelPrintDTO} after reading excel file bytes into it
	 *  
	 * @return  <code>ExcelPrintDTO</code>
	 * @throws Exception
	 */
	private static ExcelPrintDTO createExcelPrintDTO() throws Exception {

		ExcelPrintDTO excelPrintDTO = null;
		File file = null;
		FileInputStream fis = null;
		try {
			file = writeInFile();
			fis = new FileInputStream(file);
			byte[] fileBytes = IOUtils.toByteArray(fis);
			excelPrintDTO = new ExcelPrintDTO(fileBytes, file.getAbsolutePath());
		} catch (Exception e) {
			throw e;
		} finally {
			if (fis != null) {
				fis.close();
			}

			if (file != null && file.exists()) {
				file.delete();
			}			 

		}

		return excelPrintDTO;
	}

	/**
	 * Returns {@code boolean} if found headers or footers key values in {@code defaultMap}
	 * 
	 * @param defaultMap       contains headers and footers   
	 * @param header           check for headers if true
	 * @param footer           check for footers if true

	 * @return  boolean
	 */

	private static boolean ifContains(EnumMap<ExcelConstant, Object> defaultMap, boolean header, boolean footer) {
		if (header) {
			if (defaultMap.containsKey(ExcelConstant.FROM_DATE) || defaultMap.containsKey(ExcelConstant.TO_DATE)
					|| defaultMap.containsKey(ExcelConstant.REPORT_TITLE)) {
				return true;

			}
		} else if (footer) {
			if (defaultMap.containsKey(ExcelConstant.REPORTED_BY)
					|| defaultMap.containsKey(ExcelConstant.REPORTED_ON)) {
				return true;

			}
		}
		return false;
	}

	

	/**
	 *Create Report title row value taken from {@code headers}
	 * 
	 * @param headers           get report title value from it
	 * 
	 */
	private static void createReportTiltle(EnumMap<ExcelConstant, Object> headers) {
		HSSFRow reportTitleRow = excelSheetEdit.addRow();
		HSSFCell reportTitleCell = excelSheetEdit.addCell(reportTitleRow, 0, headers.get(ExcelConstant.REPORT_TITLE));
		HSSFFont reportTitleFont = getUnderLineBoldFont(22, IndexedColors.BLACK.getIndex(), "Arial",
				HSSFFont.ANSI_CHARSET, HSSFFont.U_SINGLE);
		HSSFCellStyle reportTitleStyle = workbook.createCellStyle();
		reportTitleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		reportTitleStyle.setFont(reportTitleFont);
		reportTitleCell.setCellStyle(reportTitleStyle);
		mergeRegion(reportTitleRow.getRowNum(), reportTitleRow.getRowNum(), reportTitleCell.getColumnIndex(),
				(rowListSize * 2) - 1);
		++rowNum;
	}

	

	/**
	 * Helper method to  Merger Region 
	 * 
	 * @param fromRow        merge from fromRow
	 * @param tillRow		 merge till  tillRow
	 * @param fromCell       merge fromCell 
	 * @param tillCell       merge tillCell
	 */
	@SuppressWarnings("deprecation")
	private static void mergeRegion(int fromRow, int tillRow, int fromCell, int tillCell) {
		sheet.addMergedRegion(new CellRangeAddress(fromRow, tillRow, fromCell, tillCell));
	}
	
	/**
	 * Helper method to set border on cell style
	 * 
	 * @param cellStyle   apply border on it
	 */

	private static void setBorders(HSSFCellStyle cellStyle) {
		cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
	}

	/**
	 * Helper method to create font
	 * 
	 * @param fontSize             size of font
	 * @param fontColor            color of font
	 * @param fontName             font name
	 * @param charSet              font character set
	 * 
	 * @return <code>HSSFFont</code>
	 */
	private static HSSFFont getFont(int fontSize, short fontColor, String fontName, byte charSet) {
		HSSFFont font = workbook.createFont();
		font.setCharSet(charSet);
		font.setFontHeightInPoints((short) fontSize);
		font.setColor(fontColor);
		font.setFontName(fontName);
		return font;
	}
	
	/**
	 * Helper method to create font. it calls {@code getFont} inside with additional setting of bold
	 * 
	 * @param fontSize             size of font
	 * @param fontColor            color of font
	 * @param fontName             font name
	 * @param charSet              font character set
	 * 
	 * @return <code>HSSFFont</code>
	 */
	private static HSSFFont getBoldFont(int fontSize, short fontColor, String fontName, byte charSet) {
		HSSFFont font = getFont(fontSize, fontColor, fontName, charSet);
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		return font;
	}

	/**
	 * Helper method to create font. it calls {@code getFont} inside with additional setting of underline
	 * 
	 * @param fontSize             size of font
	 * @param fontColor            color of font
	 * @param fontName             font name
	 * @param charSet              font character set
	 * @param underLine            font line
	 * 
	 * @return <code>HSSFFont</code>
	 */
	private static HSSFFont getUnderLineBoldFont(int fontSize, short fontColor, String fontName, byte charSet,
			byte underLine) {
		HSSFFont font = getBoldFont(fontSize, fontColor, fontName, charSet);
		font.setUnderline(underLine);
		return font;
	}
	
	/**
	 *  Helper method to set cell foreground color
	 *  
	 * @param cellStyle                    apply color on cell
	 * @param foregroundColor              color name
	 * @param pattern                      how to be filled with {@code foregroundColor}
	 */

	private static void setForeground(HSSFCellStyle cellStyle, short foregroundColor, short pattern) {
		cellStyle.setFillForegroundColor(foregroundColor);
		cellStyle.setFillPattern(pattern);
	}

	/**
	 * Helper method to get color from RGB Palate
	 * 
	 * @param workbook           used to get color palette from
	 * @param color              required color
	 * @param r                  dimensions of RED
	 * @param g                  dimensions of GREEN
	 * @param b                  dimensions of BLUE
	 * 
	 * @return <code>HSSFColor</code>
	 */
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

	/**
     * Returns the {@code workbook} set by {@code ExcelExportManager} 
   	 *
     * @return  <code>HSSFWorkbook</code>
     */
	public static HSSFWorkbook getWorkbook() {
		return workbook;
	}
	
	/**
     * Sets the {@code ExcelSheetEdit} that will be used by {@code ExcelExportManager}
     * to integrate custom implementation with created workbook.
   	 *
     */
	public static void setWorkbook(HSSFWorkbook workbook) {
		ExcelExportManager.workbook = workbook;
	}
	
	/**
     * Returns the {@code sheet} created on  {@code workbook} by {@code ExcelExportManager}
   	 *
     * @return  <code>HSSFSheet</code>
     */
	public static HSSFSheet getSheet() {
		return sheet;
	}

	/**
     * Sets the {@code sheet} set by {@code ExcelExportManager} 
	 *
     */
	public static void setSheet(HSSFSheet sheet) {
		ExcelExportManager.sheet = sheet;
	}


}
