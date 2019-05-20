package com.wd.api.cases;

public class Case {
private String  caseId;
private String  desc;
private String  apiId;
private String  params;
public String getCaseId() {
	return caseId;
}
public void setCaseId(String caseId) {
	this.caseId = caseId;
}
public String getDesc() {
	return desc;
}
public void setDesc(String desc) {
	this.desc = desc;
}
public String getApiId() {
	return apiId;
}
public void setApiId(String apiId) {
	this.apiId = apiId;
}
public String getParams() {
	return params;
}
public void setParams(String params) {
	this.params = params;
}
@Override
	public String toString() {
		// TODO Auto-generated method stub
		return super.toString();
	}

}
