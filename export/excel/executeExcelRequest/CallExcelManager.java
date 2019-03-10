package export.excel.executeExcelRequest;

import java.util.List;

import export.excel.common.dto.ExcelDTO;
import export.excel.common.dto.ExcelDTOManager;
import export.excel.common.dto.ExcelPrintDTO;
import export.excel.common.dto.ExcelRequestDTO;
import export.excel.common.util.ExcelExportManager;


public class CallExcelManager {


	public static ExcelPrintDTO callManager(ExcelRequestDTO excelrequestDTO) throws Exception {
	
		List<ExcelDTO> excelDTOLS=ExcelDTOManager.initialize(excelrequestDTO.getExcelDataSource()); 
		ExcelPrintDTO excelPrintDTO=null;
		if(excelrequestDTO.getExcelCustomAdd() !=null) {
			excelPrintDTO=ExcelExportManager.exportExcel(excelrequestDTO.getDefaultMap(),excelrequestDTO.getRowHeadingLS(),excelrequestDTO.getExcelCustomAdd(),excelDTOLS,excelrequestDTO.getDataIndex());
		}else {
			excelPrintDTO=ExcelExportManager.exportExcel(excelrequestDTO.getDefaultMap(),excelrequestDTO.getRowHeadingLS(),excelDTOLS,excelrequestDTO.getDataIndex());	
		}
		return excelPrintDTO;
	}

}
