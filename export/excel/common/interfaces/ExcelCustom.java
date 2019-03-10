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

package export.excel.common.interfaces;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;

/**
 * 
 * The {@code ExcelCustom} provides helper method to work on
 * Excel workbook and customize data and format it as required.
 * The {@code ExcelCustomAdd} provides implementation for these method.
 * 
 * @author Ruba Mushtaq
 * @see export.excel.common.model.ExcelCustomAdd
 * @since JDK1.5
 */
public interface ExcelCustom {
	
	/**
     * Returns the row that is added to workbook .  
     * 
     * @return  <code>HSSFRow</code>
     */
	public HSSFRow addRow();

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
	public HSSFCell addCell(HSSFRow row, int cellIndex, Object cellValue);
	
	
	/**
     * Returns the Cell that is created on {@code row} at {cellIndex} .  
     * 
     * @param    row         The row gets by {@code addRow} method where cell will be added.
     * @param   cellIndex    The cell index where cell be created.
     * 
     * @return  <code>HSSFCell</code>
     */
	public HSSFCell addCell(HSSFRow row, int cellIndex);

	/**
     * Returns the Cell where value has been placed .  
     * 
     * @param    cell       The cell which is created by calling {@code addCell} method. 
     * @param    cellValue        The cell value be placed on {@code cell}
     * 
     * @return  <code>HSSFCell</code>
     */
	public HSSFCell addCellValue(HSSFCell cell, Object cellValue);

	
	/**
     * Returns the Blank cell.  
     * 
     * @param    row         The row gets by {@code addRow} method where cell will be added.
     * @param   cellIndex    The cell index where cell be created.
     * @return  <code>HSSFCell</code>
     */
	
	public HSSFCell addBlankCell(HSSFRow row, int cellIndex);

}
