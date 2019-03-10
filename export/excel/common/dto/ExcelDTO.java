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

package export.excel.common.dto;

import java.util.List;
/**
 * 
 * The {@code ExcelDTO} contains only {@value excelDataLS} 
 * list of objects which will be set by {@code ExcelDTOManager}.
 * The class will have rows of objects that would be a data list
 * for excel rows.
 * 
 * 
 * @author Ruba Mushtaq
 * @see export.excel.common.dto.ExcelDTOManager
 * @since JDK1.5
 */

public class ExcelDTO {

    private final List<Object> excelDataLS;

	public ExcelDTO(List<Object> excelDataLS) {
		super();
		this.excelDataLS = excelDataLS;
	}
    /**
     * Return List of objects that is set by {@code ExcelDTOManager}
     * @return  <code>List<code>
     */
	public List<Object> getExcelDataLS() {
		return excelDataLS;
	}

	
}
