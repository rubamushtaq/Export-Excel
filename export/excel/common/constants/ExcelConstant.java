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

package export.excel.common.constants;

import java.util.Properties;

import export.excel.common.util.ExcelFilePropertyManager;

/**
 * Common keys for Excel Report. These values should only be placed as 
 * default headers ,footers and designated params
 
 * @author  Ruba Mushtaq
 * @since 1.5
 */
public enum ExcelConstant {
	
	/** Header title key for From Date Search   */
	FROM_DATE, 
	
	/** Header title key for To Date Search   */
	TO_DATE, 
	
	/** Footer title key for Reportee  */
	REPORTED_BY, 
	
	/** Footer title key for Reporting date  */
	REPORTED_ON, 
	
	/** Fixed title value for first sheet name */
	FIRST_SHEET, 
	
	/** Title key for first sheet name */
	REPORT_TITLE, 
	

	
	/** Title key for Search heading */
	SEARCH_TITLE;

	
	/** used to store values for all constant above*/
	private static Properties properties;

	/** required values for specific {@code ExcelConstant} constants*/
	private String value;
	

	/** load properties for the keys
	 * 
	 */
	private void init() {
		if (properties == null) {
			properties = ExcelFilePropertyManager.getProperties();
		}
		value = (String) properties.get(this.toString());
	}

	/**Returns {@code value} stores value for the {@code ExcelConstant} constant
	 * 
	 * @return <code>String<code>
	 */
	public String getValue() {
		if (value == null) {
			init();
		}
		return value;
	}
	

}
