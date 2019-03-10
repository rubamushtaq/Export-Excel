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

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.List;

import export.excel.common.annotation.ExcelRPTIgnoreFld;
import export.excel.common.interfaces.ExcelDataSource;
//put in util class
/**
 * 
 * The {@code ExcelDTOManager} is a helper class which is used to generate
 * {@code ExcelDTO} beans that will be sent to {@code ExcelExportManager} as a
 * data source for excel report generation. It won't be chosen for instantiation
 * but to create {@code ExcelDTO} objects.
 * 
 * @author Ruba Mushtaq
 * @see export.excel.common.dto.ExcelDTO
 * @see export.excel.common.interfaces.ExcelDataSource
 * @since JDK1.5
 */
public class ExcelDTOManager {

	/**
	 * Prevents to create instants of this class.
	 */
	private ExcelDTOManager() {
		
	}

	/**
	 * Return a list of {@code ExcelDTO} that will be used in report as a
	 * dataSource. It will behave as a generalized bean for excel report.
	 * It would read fields as they are arranged in polymorphic object and
	 * be set in list from count 1. 
	 * <blockquote><pre>
     *    Class implements ExcelDataSource{
     *     private int i=1;
     *     private int j=2;
     *     private int k=3;
     *    
     *    }
     *    
     *    It would be read and set as:
     *    List.add(null);
     *    List.add(1);
     *    List.add(2);
     *    List.add(3);
     * </pre></blockquote><p> This logic would help custom arrangements of fields 
     *   in Excel Report
     *
     * @param dataSourceLS
	 *            takes all Objects which implements ExcelDataSource
	 *
	 * @throws IllegalAccessException
	 *             If not able to access fields from {@value dataSourceLS}
	 *             objects
	 * 
	 * @throws IllegalArgumentException
	 *             If not able to convert {@value dataSourceLS} objects into
	 *             {@code ExcelDataSource}
	 */

	public static List<ExcelDTO> initialize(List<? extends ExcelDataSource> dataSourceLS)
			                           throws IllegalArgumentException, IllegalAccessException {
		try {
			List<ExcelDTO> dtoLs = new ArrayList<>();
			List<Object> dataSourceValLS = null;
			for (ExcelDataSource dataSource : dataSourceLS) {
				Field[] fields = dataSource.getClass().getDeclaredFields();
				dataSourceValLS = new ArrayList<>(1);
				dataSourceValLS.add(null);
				int count = 1;
				for (Field field : fields) {
					if (Modifier.isPrivate(field.getModifiers())) {
						field.setAccessible(true);
						if ( field.getName().equalsIgnoreCase("serialVersionUID")
							|| field.isAnnotationPresent(ExcelRPTIgnoreFld.class))
						continue;
						dataSourceValLS.add(count, field.get(dataSource));
						count++;

					}
				}
				ExcelDTO e = new ExcelDTO(dataSourceValLS);
				dtoLs.add(e);
			}
			return dtoLs;
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}

	}
}
