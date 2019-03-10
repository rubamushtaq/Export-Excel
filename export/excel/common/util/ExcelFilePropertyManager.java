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
 */

package export.excel.common.util;

import java.io.InputStreamReader;
import java.util.Properties;

/**
 * Property Holder class that is used to load property file of excel
 
 * @author  Ruba Mushtaq
 * @since 1.5
 */


public class ExcelFilePropertyManager {
	
	/** used to store key values pair */
	private static Properties properties;
	
	/** used to store value of a key */
	private static String value;
	
	/** path of an excel property file */
	private static final String PATH = "excel-report.properties";
	
	
	
	/** load the property file
	 * 
	 */
	private static void init() {

		properties = new Properties();
		try {
			properties.load(new InputStreamReader(
					ExcelFilePropertyManager.class.getClassLoader().getResourceAsStream(PATH), "UTF-8"));

		} catch (Exception e) {

			e.printStackTrace();
		}

	}
	
	/**Returns {@code value} stores value for {@code key}
	 * 
	 * @param  key to be searched on
	 * @return <code>String<code>
	 */

	public static String getValue(String key) {
		if (properties == null) {
			init();
		}
		return value = (String) properties.get(key);
	}

	/**Returns {@code properties} 
	 * 
	
	 * @return <code>Properties<code>
	 */
	public static Properties getProperties() {
		if (properties == null) {
			init();
		}
		return properties;
	}

	
	
}
