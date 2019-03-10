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

package export.excel.common.dto;

import java.io.Serializable;
import java.rmi.Remote;

import javax.security.auth.callback.Callback;
/**
 * 
 * The {@code ExcelPrintDTO} is a final thread safe class contains 
 * information of generated Excel file and would return back to calling
 * code and takes the use of as it requires.
 * 
 * @author Ruba Mushtaq
 * @since  JDK1.5
 */
public final class ExcelPrintDTO  implements Serializable,Callback,Remote{

	/** use serialVersionUID from JDK 1.0.2 for interoperability */
	private static final long serialVersionUID = 1L;
	
	/** The value is used for excel file byte storage */
	private final byte[] excelFileBytes;
	
	/** The value is used for excel file path storage */
	private final String excelFilePath;

	/**
	 * Initializes a newly created {@code ExcelPrintDTO} that represents
	 * excel file bytes and its path.
	 * 
	 * @param excelFileBytes
	 * @param excelFilePath
	 */
	public ExcelPrintDTO(byte[] excelFileBytes,String excelFilePath){
		this.excelFileBytes=excelFileBytes;
		this.excelFilePath=excelFilePath;
	}
	
	/**
	 * Returns a {@value excelFileBytes}.
	 */
	public byte[] getExcelFileBytes() {
		return excelFileBytes;
	}

	/**
	 * Returns a {@value excelFilePath}.
	 */
	public String getExcelFilePath() {
		return excelFilePath;
	}
	
	
	
	
}
