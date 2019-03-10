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
 *
 *
 */
package export.excel.common.model;

import java.io.Serializable;
import java.rmi.Remote;
import javax.security.auth.callback.Callback;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import export.excel.common.interfaces.ExcelCustom;

/**
 * 
 * The {@code ExcelSheetEdit} is helper class provides implementation of common 
 * helper method to work on Excel workbook and customize data and format it as required.
 * The {@code ExcelCustomAdd} provides implementation for these method.
 * 
 * @author Ruba Mushtaq
 * @see export.excel.common.model.ExcelCustomAdd
 * @since JDK1.5
 */
public final class ExcelSheetEdit implements ExcelCustom,Callback, Serializable, Remote{
	
	/** use serialVersionUID from JDK 1.0.2 for interoperability */
	private static final long serialVersionUID = 6400127106398498123L;
	
	/** The value is used for excel workbook storage */
	transient  private final HSSFWorkbook workbook;
	
	/** The value is used for excel worksheet storage in {@value workbook} */
	transient  private final HSSFSheet sheet;
	
	/** The value is used to maintain row count in {@code workbook} */
	transient  private Integer rowCount;
	
	/**
	 * Initialized {@code ExcelSheetEdit}
	 * 
	 * @param workbook
	 * @param sheet
	 * @param rowCount
	 */
	public ExcelSheetEdit(HSSFWorkbook workbook, HSSFSheet sheet, Integer rowCount) {
		this.workbook = workbook;
		this.sheet = sheet;
		this.rowCount = rowCount;
	
		
	}
	
	private  ExcelSheetEdit(){
		workbook=null;
		sheet=null;
		rowCount=null;
		
		
	}

	public void setRowCount(Integer rowCount) {
		this.rowCount = rowCount;
	}

	/**
     * Returns the row that is added to workbook .  
     * 
     * @return  <code>HSSFRow</code>
     */
	public final HSSFRow addRow() {
		HSSFRow headerRow = sheet.createRow(++rowCount);
		return headerRow;
	}

	/**
     * Returns the Cell that is added to the row with value .  
     * <blockquote><pre>
     *  	HSSFRow row=this.addRow();
     *      HSSFCell cell=this.addCell(row,0,"123") - 1st cell creates at the row
     *      
     *  </pre></blockquote>
     * @param    row         The row gets by {@code addRow} method where cell will be added.
     * @param    cellIndex   The cell index where cell be created.
     * @param    cellValue   The cell value be placed on new cell created at {@code cellIndex}} 
     * 
     * @return  <code>HSSFCell</code>
     */
	public final HSSFCell addCell(HSSFRow row, int cellIndex, Object cellValue) {
		HSSFCell cell = row.createCell(cellIndex);
		HSSFCellStyle rowHeadingStyle = workbook.createCellStyle();
		rowHeadingStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		cell.setCellValue("" + cellValue);
		return cell;
	}

	/**
     * Returns the Cell where value has been placed .  
     * 
     * @param    cell       The cell which is created by calling {@code addCell} method. 
     * @param    cellValue        The cell value be placed on {@code cell}
     * @return  <code>HSSFCell</code>
     */
	
	public final HSSFCell addCellValue(HSSFCell cell, Object cellValue) {
		cell.setCellValue("" + cellValue);
		return cell;
	}
	
	/**
     * Returns the Cell that is created on {@code row} at {cellIndex} .  
     * 
     * @param    row         The row gets by {@code addRow} method where cell will be added.
     * @param   cellIndex    The cell index where cell be created.
     * 
     * @return  <code>HSSFCell</code>
     */
	public final HSSFCell addCell(HSSFRow row, int cellIndex) {
		HSSFCell cell = row.createCell(cellIndex);
		return cell;
	}
	
	/**
     * Returns the Blank cell.  
     * 
     * @param    row         The row gets by {@code addRow} method where cell will be added.
     * @param   cellIndex    The cell index where cell be created.
     * @return  <code>HSSFCell</code>
     */
	
	public final HSSFCell addBlankCell(HSSFRow row, int cellIndex) {
		HSSFCell cell = row.createCell(cellIndex);
		return cell;
	}
	

	/**
     * Returns the count of last added row number {@code rowCount }in the sheet.  
     *
     * @return  <code>Integer</code>
     */

	public Integer getRowCount() {
		return rowCount;
	}
	
	/**
     * Returns the {@code workbook} set by {@code ExcelCustomAdd} 
   	 *
     * @return  <code>HSSFWorkbook</code>
     */

	public HSSFWorkbook getWorkbook() {
		return workbook;
	}
	
	/**
     * Returns the {@code sheet} set by {@code ExcelCustomAdd} 
   	 *
     * @return  <code>HSSFSheet</code>
     */

	public HSSFSheet getSheet() {
		return sheet;
	}


	
	
	

}
