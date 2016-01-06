package excel;

public class MessageDTO
{
	/**
	 * 文件名
	 */
	private String fileName = "";
	
	/**
	 * 字第
	 */
	private String number = "";
	
	/**
	 * 原告
	 */
	private String palinTiff = "";
	
	/**
	 * 被告
	 */
	private String defendant = "";
	
	/**
	 * 案由
	 */
	private String caseReson = "";
	
	/**
	 * 结案方式
	 */
	private String closeWay = "";

	/**
	 * 审判员
	 */
	private String judge = "";
	
	/**
	 * 审判长
	 */
	private String chiefJudge = "";
	
	/**
	 * 立案日期
	 */
	private String recordDate;
	
	/**
	 * 结案日期
	 */
	private String endDate;
	
	/**
	 * 合议庭成员
	 */
	private String collegiateMembers;
	
	
	public String getFileName()
	{
		return fileName;
	}

	public void setFileName(String fileName)
	{
		this.fileName = fileName;
	}

	public String getNumber()
	{
		return number;
	}

	public void setNumber(String number)
	{
		this.number = number;
	}

	public String getPalinTiff()
	{
		return palinTiff;
	}

	public void setPalinTiff(String palinTiff)
	{
		this.palinTiff = palinTiff;
	}

	public String getDefendant()
	{
		return defendant;
	}

	public void setDefendant(String defendant)
	{
		this.defendant = defendant;
	}

	public String getCaseReson()
	{
		return caseReson;
	}

	public void setCaseReson(String caseReson)
	{
		this.caseReson = caseReson;
	}

	public String getCloseWay() {
		return closeWay;
	}

	public void setCloseWay(String closeWay) {
		this.closeWay = closeWay;
	}

	public String getJudge() {
		return judge;
	}

	public void setJudge(String judge) {
		this.judge = judge;
	}

	public String getChiefJudge() {
		return chiefJudge;
	}

	public void setChiefJudge(String chiefJudge) {
		this.chiefJudge = chiefJudge;
	}

	public String getRecordDate() {
		return recordDate;
	}

	public void setRecordDate(String recordDate) {
		this.recordDate = recordDate;
	}

	public String getEndDate() {
		return endDate;
	}

	public void setEndDate(String endDate) {
		this.endDate = endDate;
	}

	public String getCollegiateMembers() {
		return collegiateMembers;
	}

	public void setCollegiateMembers(String collegiateMembers) {
		this.collegiateMembers = collegiateMembers;
	}
}
