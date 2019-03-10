package export.excel.executeExcelRequest;

import java.util.ArrayList;
import java.util.Date;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import export.excel.common.constants.ExcelConstant;
import export.excel.common.dto.ExcelPrintDTO;
import export.excel.common.dto.ExcelRequestDTO;
import export.excel.common.dto.UserDTO;
import export.excel.common.interfaces.ExcelDataSource;

public class RunExcelWrite {

	public static void main(String[] args) throws Exception {
		List<String> columns = new ArrayList<>();
		columns.add("Column1");
		columns.add("Column2");
		columns.add("Column3");
		columns.add("Column4");
		columns.add("Column5");
		columns.add("Column6");
		columns.add("Column7");

		List<ExcelDataSource> userList = new ArrayList<>();
		UserDTO user1 = new UserDTO();
		user1.setUserID(1);
		user1.setUserName("Test1");
		user1.setI(11);
		user1.setJ(112);
		user1.setNullSecond("232");
		user1.setCommitteeId(111);
		user1.setCommitteeName("Committee1");

		UserDTO user2 = new UserDTO();
		user2.setUserID(2);
		user2.setUserName("Test2");
		user2.setI(22);
		user2.setJ(212);
		user2.setNullSecond("2323");
		user2.setCommitteeId(222);
		user2.setCommitteeName("Committee2");

		UserDTO user3 = new UserDTO();
		user3.setUserID(3);
		user3.setUserName("Test3");
		user3.setI(33);
		user3.setJ(312);
		user3.setNullSecond(null);
		user3.setCommitteeId(333);
		user3.setCommitteeName("Committee3");

		userList.add(user1);
		userList.add(user2);
		userList.add(user3);
		int[] indexPosition = new int[7];
		indexPosition[0] = 5; //
		indexPosition[1] = 4;
		indexPosition[2] = 6;
		indexPosition[3] = 3;
		indexPosition[4] = 2;
		indexPosition[5] = 7;
		indexPosition[6] = 1;

		LinkedHashMap<String, Object> map1 = new LinkedHashMap<>();
		map1.put("fromDate", new Date());
		map1.put("toDate", new Date());

		map1.put("exportedBy", 12);
		map1.put("exportedOn", new Date());

		Map<String, Object> map2 = new HashMap<>();
		map2.put("Date1", new Date());
		map2.put("Date2", new Date());

		EnumMap<ExcelConstant, Object> map3 = new EnumMap<>(ExcelConstant.class);
		map3.put(ExcelConstant.FROM_DATE, "2013-34-43");
		map3.put(ExcelConstant.TO_DATE, "2011-02-34");
		map3.put(ExcelConstant.REPORTED_BY, "Admin (234)");
		map3.put(ExcelConstant.REPORTED_ON, "2018-29-03 45:60:34");
		map3.put(ExcelConstant.REPORT_TITLE, "Report Title");

		Map<String, Object> customParams = new HashMap<>();
		customParams.put("param1", "Param1");
		customParams.put("param2", "Param2");

		customParams.put("columnSize", columns.size());
		ExcelRequestDTO excelrequestDTO = new ExcelRequestDTO(map3, columns, userList, indexPosition);
		@SuppressWarnings("unused")
		ExcelPrintDTO dto = CallExcelManager.callManager(excelrequestDTO);
		// read dto as required

	}

}
