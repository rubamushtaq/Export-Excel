package export.excel.common.dto;

import export.excel.common.annotation.ExcelRPTIgnoreFld;
import export.excel.common.interfaces.ExcelDataSource;

public class UserDTO implements ExcelDataSource{

	//@ExcelRPTIgnoreFld
	private int userID;//0  // 2
	private String userName;//1 //1
	private int i;//2    //0
	//@ExcelRPTIgnoreFld
	private String  FirstSecond; //0//skip

	private int j;

	private String committeeName;
	//@ExcelRPTIgnoreFld
	private String  nullSecond;//4//skip
	@ExcelRPTIgnoreFld 
	private int committeeId;
	@ExcelRPTIgnoreFld 
	private String empName;
	@ExcelRPTIgnoreFld 
	private long empNo;
	
	
	/*
	 * public int getJ() { return j; } public void setJ(int j) { this.j = j; }
	 */
	public int getUserID() {
		return userID;
	}
	public void setUserID(int userID) {
		this.userID = userID;
	}
	public String getUserName() {
		return userName;
	}
	public void setUserName(String userName) {
		this.userName = userName;
	}
	public int getI() {
		return i;
	}
	public void setI(int i) {
		this.i = i;
	}
	public String getFirstSecond() {
		return FirstSecond;
	}
	public void setFirstSecond(String firstSecond) {
		FirstSecond = firstSecond;
	}
	public int getJ() {
		return j;
	}
	public void setJ(int j) {
		this.j = j;
	}
	public String getNullSecond() {
		return nullSecond;
	}
	public void setNullSecond(String nullSecond) {
		this.nullSecond = nullSecond;
	}
	public String getCommitteeName() {
		return committeeName;
	}
	public void setCommitteeName(String committeeName) {
		this.committeeName = committeeName;
	}
	public int getCommitteeId() {
		return committeeId;
	}
	public void setCommitteeId(int committeeId) {
		this.committeeId = committeeId;
	}
	public String getEmpName() {
		return empName;
	}
	public void setEmpName(String empName) {
		this.empName = empName;
	}
	public long getEmpNo() {
		return empNo;
	}
	public void setEmpNo(long empNo) {
		this.empNo = empNo;
	}
	
	
	
}
