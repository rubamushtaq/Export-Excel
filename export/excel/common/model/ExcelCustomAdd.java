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
 *
 */
package export.excel.common.model;

import java.io.Serializable;
import java.rmi.Remote;
import java.util.Map;
import javax.security.auth.callback.Callback;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import export.excel.common.interfaces.ExcelCustom;
import export.excel.common.util.ExcelReportManager;

/**
 * 
 * The {@code ExcelCustomAdd} is an Abstract class represents
 * {@code HSSFWorkbook} {@value workbook} and {@code HSSFSheet} {@code sheet}
 * with default implementation of general functionality and custom
 * implementation of Header and footer via {@code ExcelSheetEdit}
 * {@code excelSheetEdit} maintained by {@code rowNum}to be added in Excel.
 * Every child needs to implement these behavior to define custom implementation
 * for an excel. {@code ExcelCustomAdd} child classes cannot override default
 * implementation though can provide custom implementation with default
 * behavior.
 * 
 * @author Ruba Mushtaq
 * @see org.apache.poi.hssf.usermodel.HSSFWorkbook
 * @see org.apache.poi.hssf.usermodel.HSSFSheet
 * @see export.excel.common.model.ExcelSheetEdit
 * @since JDK1.5
 */

public abstract class ExcelCustomAdd implements ExcelCustom,Callback, Serializable, Remote {

	/** use serialVersionUID from JDK 1.0.2 for interoperability */
	private static final long serialVersionUID = -1506863074641818256L;

	/** The value is used for excel workbook storage */
	transient private final HSSFWorkbook workbook;

	/** The value is used for excel work sheet storage in {@code workbook} */
	transient private final HSSFSheet sheet;

	/** The value is used to maintain rows in {@code workbook} */
	transient private int rowNum;

	/** {@code ExcelCustomAdd } delegates the implementation to {@code excelSheetEdit}.
	 *  The value is used to wrap{@value workbook} {@code sheet} {@code rowNum}
	 *  It will provide concrete implementation of working of {@code workbook}
	 *  @{@code sheet} */
	transient private  ExcelSheetEdit excelSheetEdit;
	
	/** The value is used to store custom headers */
	protected Map<String, Object> customHeaders;
	
	/** The value is used to store custom footers */
	protected Map<String, Object> customFooters;
	
	/** The value is used to store custom params */
	protected Map<String, Object> customParams;
	

	/**
	 * Initializes a newly created {@code ExcelCustomAdd} object so that it
	 * represents {@value workbook} {@value sheet} {@value rowNum}
	 * {@value excelSheetEdit}. The use of this constructor is to create an object
	 * with default {@code ExcelSheetEdit} {@value excelSheetEdit} implementation to
	 * provide custom implementation on {@value workbook}.
	 */
		
	protected ExcelCustomAdd(Map<String, Object> customHeaders,Map<String, Object> customFooters,Map<String, Object> customParams) {
		this.workbook = ExcelReportManager.getWorkbook();
		this.sheet = ExcelReportManager.getSheet();
		this.rowNum = ExcelReportManager.getRowNum();
		excelSheetEdit = new ExcelSheetEdit(workbook, sheet, rowNum);
		this.customHeaders=customHeaders;
		this.customFooters=customFooters;
		this.customParams=customParams;

	}

	/**
	 * This will be used add custom headers and style it into workbook 
	 * as required by calling helper methods. Here are some examples of
	 * how it can be used: <blockquote>
	 * 
	 * <blockquote><pre>
	 * String cH=customHeaders.get("key");
	 * Row row=this.addRow();
	 * this.addCell(row,0,"val");
	 * </pre></blockquote>
	 */
	public void customHeaderImpl() {

	}
	
	/**
	 * This will be used add custom footer and style it into workbook 
	 * as required by calling helper methods. Here are some examples of
	 * how it can be used: <blockquote>
	 * 
	 * <blockquote><pre>
	 * String cH=customFooters.get("key");
	 * Row row=this.addRow();
	 * this.addCell(row,0,"val");
	 * </pre></blockquote>
	 */
	public void customFooterImpl() {

	}

	/**
	 * This will be used add custom params and style it into workbook 
	 * as required by calling helper methods. Here are some examples of
	 * how it can be used: <blockquote>
	 * 
	 * <blockquote><pre>
	 * String cH=customParams.get("key");
	 * Row row=this.addRow();
	 * this.addCell(row,0,"val");
	 * </pre></blockquote>
	 */
	public void customParams() {

	}
	
	/**
     * Returns the row that is added to workbook .  
     * 
     * @return  <code>HSSFRow</code>
     */
	public final HSSFRow addRow() {
		return excelSheetEdit.addRow();
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
	public HSSFCell addCell(HSSFRow row, int cellIndex, Object cellValue) {
		return excelSheetEdit.addCell(row, cellIndex, cellValue);
	}

	/**
     * Returns the Cell where value has been placed .  
     * 
     * @param    cell             The cell which is created by calling {@code addCell} method. 
     * @param    cellValue        The cell value be placed on {@code cell}
     * @return  <code>HSSFCell</code>
     */
	public final HSSFCell addCellValue(HSSFCell cell, Object cellValue) {
		return excelSheetEdit.addCellValue(cell, cellValue);
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
		return excelSheetEdit.addCell(row, cellIndex);
	}

	/**
     * Returns the Blank cell.  
     * 
     * @param    row         The row gets by {@code addRow} method where cell will be added.
     * @param   cellIndex    The cell index where cell be created.
     * @return  <code>HSSFCell</code>
     */
	public HSSFCell addBlankCell(HSSFRow row, int cellIndex) {
		return excelSheetEdit.addBlankCell(row, cellIndex);
	}

	/**
     * Returns the {@code ExcelSheetEdit} that will be used to get composed objects
     * {@code HSSFWorkbook} and {@code HSSFSheet} to format rows and cells  
   
     * @return  <code>ExcelSheetEdit</code>
     */
	public ExcelSheetEdit getExcelSheetEdit() {
		return excelSheetEdit;
	}
	
	/**
     * Sets the {@code ExcelSheetEdit} that will be used by {@code ExcelExportManager}
     * to integrate custom implementation with created workbook.
   
     * @return  <code>ExcelSheetEdit</code>
     */
	public void setExcelSheetEdit(ExcelSheetEdit excelSheetEdit) {
		this.excelSheetEdit = excelSheetEdit;
	}

	/**
     * Returns the {@code Map} that stores {@code customParams} 
   
     * @return  <code>Map</code>
     */
	public Map<String, Object> getCustomParams() {
		return customParams;
	}

	
	public void setCustomParams(Map<String, Object> customParams) {
		this.customParams = customParams;
	}

	/**
     * Returns the {@code Map} that stores {@code customHeaders} 
   
     * @return  <code>Map</code>
     */
	public Map<String, Object> getCustomHeaders() {
		return customHeaders;
	}

	public void setCustomHeaders(Map<String, Object> customHeaders) {
		this.customHeaders = customHeaders;
	}

	public Map<String, Object> getCustomFooters() {
		return customFooters;
	}

	/**
     * Returns the {@code Map} that stores {@code customFooters} 
   
     * @return  <code>Map</code>
     */
	public void setCustomFooters(Map<String, Object> customFooters) {
		this.customFooters = customFooters;
	}
	
	

}
