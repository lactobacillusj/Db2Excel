package com.lactovacillusj.util;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Tools {
	public Connection connectDB(String driver, String url, String id, String pw) throws Exception {
		Connection conn = null;
		Class.forName(driver);
		conn = DriverManager.getConnection(url, id, pw);
		return conn;
	}

	public List<String> getTableList(Connection conn) throws Exception{
		List<String> tableList = new ArrayList<String>();
		
		String sql = "SELECT TABLE_NAME FROM USER_TABLES";
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery(sql);
		while(rs.next()) 
			tableList.add(rs.getString(1));
			
		return tableList;
	}
	
	public List<?> db2List(Connection conn, String tableName) throws Exception{
		List<Object> tableInfo = new ArrayList<>();
		String sql = "SELECT * FROM " + tableName;
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery(sql);
		ResultSetMetaData rsmd = rs.getMetaData();
		
		int columncnt = rsmd.getColumnCount();
		String[] title = new String[columncnt];
		
		for(int i = 0 ; i<title.length ; i++) {
			title[i] = rsmd.getColumnName(i+1);
		}
		tableInfo.add(title); //title 배열에 첫 위치만 tableInfo의 첫 리스트에 담겨있음
		
		while(rs.next()) {
			Map<String, String> tableRow = new HashMap<String, String>();
			for(int i = 0 ; i<title.length ; i++) {
				String key = title[i];
				String value = rs.getString(key);
				tableRow.put(key,value);
			}
			tableInfo.add(tableRow);
		}
		return tableInfo;
		
	}
	
	public void putTableInfo(XSSFWorkbook wb, String tableName, List tableInfo) {
		XSSFSheet sheet = wb.createSheet(tableName);
		XSSFRow row = null;
		XSSFCell cell = null;
		
		for(int i = 0 ; i < tableInfo.size(); i++) {
			row = sheet.createRow(i);
			
			String[] title = (String[]) tableInfo.get(0);
			
			if(i==0) {
				for(int j = 0 ; j < title.length ; j++) {
					cell = row.createCell(j);
					cell.setCellValue(title[j]);
				}
			}else {
				Map<String,String> dataMap = (Map) tableInfo.get(i);
				
				for(int j = 0 ; j < title.length ; j++) {
					cell = row.createCell(j);
					//데이터베이스에 데이터들에서 컬럼 값을 키로 각 행의 데이터를 가져와서 셀에 넣는다.
					cell.setCellValue(dataMap.get(title[j]));
				}
			}
		}
	}
	
	public void writeExcel(XSSFWorkbook wb, String fileName) throws Exception{
		FileOutputStream fos = new FileOutputStream(fileName);
		
		wb.write(fos);
	}
	
}
