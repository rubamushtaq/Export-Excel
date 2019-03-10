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

package export.excel.common.util;

import java.io.Serializable;
import java.rmi.Remote;
import javax.security.auth.callback.Callback;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import export.excel.common.constants.ExcelConstant;

/**
 * 
 * The {@code ExcelReportManager} is helper class comprises {@code HSSFWorkbook},
 * {@code HSSFSheet} to work. It should not be called directly as
 * {@code ExcelCustomAdd} would be calling inside providing methods to work 
 *  on the them.
 * 
 * @author Ruba Mushtaq
 * @see export.excel.common.model.ExcelCustomAdd
 * @since JDK1.5
 */
public class ExcelReportManager implements Callback, Serializable, Remote{

	/** use serialVersionUID from JDK 1.0.2 for interoperability */
	private static final long serialVersionUID = 3907510841686901604L;
	
	/** The value is used to maintain rows in {@code workbook} */
	 transient public  static Integer rowNum = null;
	 
	/** The value is used for excel workbook storage */
	 transient private static HSSFWorkbook workbook;
	 
	 /** The value is used for excel worksheet storage in {@code workbook} */
	 transient private static HSSFSheet sheet;

	private ExcelReportManager(){
		
	}
	/**
     * Returns the {@code Integer} stores Rows count in Excel.
   
     * @return  <code>Integer</code>
     */
	public static Integer getRowNum() {
		return rowNum = -1;
	}

	/**
     * Returns the newly created {@code workbook}
   
     * @return  <code>HSSFWorkbook</code>
     */
	public static HSSFWorkbook getWorkbook() {
		workbook = new HSSFWorkbook();
		return workbook;
	}

	/**
     * Returns the newly created sheet {@code sheet} from {@code workbook} .
   
     * @return  <code>HSSFSheet</code>
     */
	public static HSSFSheet getSheet() {
		sheet = workbook.createSheet(ExcelConstant.FIRST_SHEET.getValue());
		//sheet.setRightToLeft(true); 
		return sheet;
	}

}
