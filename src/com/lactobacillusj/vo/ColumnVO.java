package com.lactobacillusj.vo;

public class ColumnVO {
	private String name;
	private String type;
	private int maxLength;
	
	public ColumnVO(String name) {
		this.name = name;
	}
	
	public ColumnVO(String name, String type, int maxLength) {
		this.name = name;
		this.type = type;
		this.maxLength = maxLength;
	}
	
	public String getName() {
		return name;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public int getMaxLength() {
		return maxLength;
	}
	public void setMaxLength(int maxLength) {
		this.maxLength = maxLength;
	}
}
