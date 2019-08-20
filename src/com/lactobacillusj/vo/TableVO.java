package com.lactobacillusj.vo;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import com.lactovacillusj.util.Tools_upgrade;

public class TableVO {
	private String name;
	private List<ColumnVO> column = new ArrayList<ColumnVO>();
	private List<Map<String,String>> data = new ArrayList<Map<String,String>>();
	
	public TableVO(String name) {
		this.name = name;
	}
	
	public String getName() {
		return name;
	}
	
	public List<Map<String,String>> getData(){
		return data;
	}
	public void setData(List<Map<String, String>> data) {
		this.data = data;
	}
	public List<ColumnVO> getColumn(){
		return column;
	}
	public void addColumn(ColumnVO column) {
		this.column.add(column);
	}
	public List<String> getColumnNameList(){
		List<String> columnNameList = new ArrayList<String>();
		
		for(int i = 0 ; i < column.size(); i++) {
			columnNameList.add(column.get(i).getName());
		}
		return columnNameList;
	}
	public void loadColumn(Connection conn) throws Exception{
		this.column = new ArrayList<ColumnVO>();
		
		//컬럼이름, 데이터타입, 데이터길이만 메타데이터로 사용 
		String sql = "SELECT column_name, data_type, data_length FROM USER_TAB_COLUMNS"
				+ " WHERE 1=1"
				+ " AND TABLE_NAME = '" + name + "'"
				+ " ORDER BY COLUMN_ID";
		
		Statement stmt = conn.createStatement();
		
		ResultSet rs = stmt.executeQuery(sql);
		
		while(rs.next()) {
			ColumnVO column = new ColumnVO(rs.getString("column_name"));
			column.setType(rs.getString("data_type"));
			column.setMaxLength(rs.getInt("data_length"));
			
			addColumn(column);
		}
		
	}
	public void loadData(Connection conn) throws Exception{
		this.data = new ArrayList<Map<String,String>>();
		String sql = "SELECT * FROM " + name;
		this.data = Tools_upgrade.getListMap(conn, sql);
	}
}
