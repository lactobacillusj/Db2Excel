package com.lactovacillusj.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
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

import com.lactobacillusj.vo.ColumnVO;
import com.lactobacillusj.vo.TableVO;

public class Tools_upgrade {
	
	public static Connection connectDB(String driver, String url, String id, String pwd) throws Exception{

		Connection conn = null;

		Class.forName(driver);
		conn = DriverManager.getConnection(url, id, pwd);

		System.out.println("========== connectDB success ==========");
		System.out.println("   driver : " + driver);
		System.out.println("   url : " + url);
		System.out.println("   id : " + id);
		System.out.println();

		return conn;
	}
	
	public static List<String> getListFirstColumn(Connection conn, String sql) throws Exception{
		
		List<String> list = new ArrayList<String>();
		
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery(sql);
		
		while(rs.next()){
			list.add(rs.getString(1));
		}

		System.out.println("========== getDBdataFirstColumn ==========");
		System.out.println("   "+sql);
		System.out.println();

		
		return list;
	}

	public static List<Map<String, String>> getListMap(Connection conn,	String sql) throws Exception {
		
		List<Map<String, String>> list = new ArrayList<Map<String, String>>();
		
		Statement stmt = conn.createStatement();
		ResultSet rs = stmt.executeQuery(sql);
		
		ResultSetMetaData rsmd = rs.getMetaData();
		int columnCnt = rsmd.getColumnCount();
		
		while(rs.next()){
			Map<String, String> row = new HashMap<String,String>();
			
			for(int i=0; i<columnCnt; i++){
				String columnName = rsmd.getColumnName(i+1);
				row.put(columnName, rs.getString(columnName));
			}

			list.add(row);
		}
		
		System.out.println("========== getDBdataAll ==========");
		System.out.println("   "+sql);
		System.out.println();

		
		return list;
	}
	
	public static void createTable(Connection conn, TableVO table) throws Exception{

		String tableName = table.getName();
		List<ColumnVO> columnList = table.getColumn();
		
		Statement stmt = conn.createStatement();
		ResultSet rs = null;
		
		String sql = "CREATE TABLE " + tableName + "( ";
		for(int i=0; i<columnList.size(); i++){
			ColumnVO column = columnList.get(i);
			String columnName = column.getName();
			String columnType = column.getType();
			int maxLength = column.getMaxLength();
			
			if(i!=0)
				sql += ",";

			sql += columnName + " " + columnType;
			if(!columnType.equalsIgnoreCase("date"))
					sql += "(" + maxLength + ")";
			
		}
		
		sql += ")";

		System.out.println("========== createTable ==========");
		System.out.println("   " + sql);
		try {
			rs = stmt.executeQuery(sql);
		} catch (Exception e) {
			if(e.getMessage().indexOf("ORA-00955")!=-1)
				System.out.println("   !!![skip]already created table : " + tableName);
			else
				e.printStackTrace();
		}
		
		System.out.println("   creat success");	
		System.out.println();
		

	}

	public static void insertData(Connection conn, TableVO table) throws Exception{
		
		List<Map<String,String>> data = table.getData();
		List<String> columnList = table.getColumnNameList();
		List<ColumnVO> column = table.getColumn();
		
		String sql = "INSERT INTO " + table.getName() + "(";
		String sql2 = ") VALUES(";
		
		for(int i=0; i<columnList.size(); i++){

			if(i!=0){ sql += ","; sql2 +=","; } 
			
			sql += columnList.get(i);
			sql2 += "?";
			
		}
		sql = sql+sql2 + ")";

		PreparedStatement pstmt = conn.prepareStatement(sql);
		
		for(int i=0; i<data.size(); i++){
			System.out.println("========== insertData ==========");
			System.out.println("   " + sql);

			Map<String,String> rowMap = data.get(i);
			
			for(int j=0; j<columnList.size(); j++){
				setPstmtValue(j, pstmt, column, rowMap);				
			}

			
			int result = pstmt.executeUpdate();
			
			System.out.println("   insert success");	
			System.out.println();
		}
		 
		
	}

	private static void setPstmtValue(int idx, PreparedStatement pstmt, List<ColumnVO> column, Map<String, String> row) throws Exception {
		String columnName = column.get(idx).getName();
		String columnType = column.get(idx).getType();
		
		String value = row.get(columnName);

		int pstmtIdx = idx +1;
		
		switch (columnType) {
		case "VARCHAR2" :
			pstmt.setString(pstmtIdx, value);
			break;
		case "NUMBER" :
			if(!value.equals(""))
				pstmt.setInt(pstmtIdx, Integer.valueOf(value));
			else 
				pstmt.setString(pstmtIdx, value);
			break;
		case "DATE" : // 에러
			if(!value.equals("")){
				pstmt.setDate(pstmtIdx, Date.valueOf(value.substring(0,10)));
			}
			else 
				pstmt.setString(pstmtIdx, value);

			break;			
		}
		
		
	}
	
	public static void writeExcel(XSSFWorkbook wb, String fileName) throws Exception{

		FileOutputStream fos = new FileOutputStream(fileName);
		
		wb.write(fos);

		System.out.println("========== writeExcel ==========");
		System.out.println("   "+ fileName);
		System.out.println();
		
		//wb에 저장된 데이터를 지정한 경로에 출력한다.
	}
	
	public static XSSFSheet setTableInfoSheet(XSSFWorkbook wb){
		
		XSSFSheet sheet = wb.createSheet("tableInfo");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = null;
		
		cell = row.createCell(0);
		cell.setCellValue("tableName");
		cell = row.createCell(1);
		cell.setCellValue("columnName");
		cell = row.createCell(2);
		cell.setCellValue("columnType");
		cell = row.createCell(3);
		cell.setCellValue("maxLength");

		System.out.println("========== setTableInfoSheet ==========");
		System.out.println("setTableInfoSheet_Table_Column_name_set");
		System.out.println();
		//첫 행에 직접적은 컬럼 이름을 출력한다. 
		
		return sheet;
	}
	
	public static int addTableInfoRow(XSSFSheet sheet, int idx, TableVO table){

		XSSFCell cell = null;
		
		List<ColumnVO> columnList = table.getColumn();
		
		for(int i=0; i<columnList.size(); i++){
			
			XSSFRow row = sheet.createRow(idx++);
			ColumnVO column = columnList.get(i);
			
			cell = row.createCell(0);
			cell.setCellValue(table.getName());
			cell = row.createCell(1);
			cell.setCellValue(column.getName());
			cell = row.createCell(2);
			cell.setCellValue(column.getType());
			cell = row.createCell(3);
			cell.setCellValue(column.getMaxLength());
			
		}
		
		System.out.println("========== addTableInfoRow ==========");
		System.out.println("   "+ table.getName());
		System.out.println();
		//ColumnVO 형에 맞는 column 정보를 가져와 sheet에 column 밑으로 table 데이터를 모두 넣는다. 
		return idx;
	}
	
	public static void addVOToSheet(XSSFWorkbook wb, TableVO table){
		
		XSSFSheet sheet = wb.createSheet(table.getName());
		
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = null;
		
		List<String> columnList = table.getColumnNameList();
		
		for(int i=0; i<columnList.size(); i++){
			cell = row.createCell(i);
			cell.setCellValue(columnList.get(i));			
		}
		
		List<Map<String,String>> data = table.getData();

		for(int i=0; i<data.size(); i++){
			
			row = sheet.createRow(i+1);
			Map<String, String> dataMap = data.get(i);
						
			for(int j=0; j<columnList.size(); j++){
				cell = row.createCell(j);
				cell.setCellValue(dataMap.get(columnList.get(j)));
			}
		}

		System.out.println("========== addVOToSheet ==========");
		System.out.println("   "+ table.getName());
		System.out.println();
		//테이블에 저장된 하나의 테이블 이름 별 데이터를 시트에 넣음
		
	}

	public static List<TableVO> excelToVOList(String fileName) throws Exception{

		System.out.println("========== excelToVO ==========");

		List<TableVO> list = new ArrayList<TableVO>();
		
		FileInputStream fis = new FileInputStream(fileName);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		readTableList(wb, list);

		
		for(int i=0; i<list.size(); i++){
			readTableData(wb, list.get(i));
		}
		
		System.out.println();
		//readTableList, readTableData를 이용하여 테이블 정보를 가져온다.
		
		//readTableList
		//readTableData tableVO형으로 가져온 table에 데이터를 넣어준다. 
		return list;
	}

	
	public static void readTableList(XSSFWorkbook wb, List<TableVO> list) throws Exception{

		XSSFSheet sheetTableInfo = wb.getSheet("tableInfo");

		int firstRowNo = sheetTableInfo.getFirstRowNum();
		int lastRowNo = sheetTableInfo.getLastRowNum();

		TableVO table = new TableVO("");

		for(int i=firstRowNo+1; i<lastRowNo+1; i++){
			
			XSSFRow row = sheetTableInfo.getRow(i);
			
			String tableName = row.getCell(0).getStringCellValue();
			String columnName = row.getCell(1).getStringCellValue();
			String columnType = row.getCell(2).getStringCellValue();
			int maxLength = (int) row.getCell(3).getNumericCellValue();
			
			if(!tableName.equals(table.getName())){
				table= new TableVO(tableName);
				list.add(table);
				System.out.println("   "+ table.getName());
			}
			ColumnVO column = new ColumnVO(columnName, columnType, maxLength);
			
			table.addColumn(column);
			
		}
		
	}
	
	public static void readTableData(XSSFWorkbook wb, TableVO table){
		
		XSSFSheet sheet = wb.getSheet(table.getName());
		XSSFRow titleRow = sheet.getRow(0);

		int firstRowNo = sheet.getFirstRowNum();
		int lastRowNo = sheet.getLastRowNum();

		List<Map<String,String>> data = new ArrayList<Map<String,String>>();
		
		for(int i=firstRowNo+1; i<lastRowNo+1; i++){

			Map<String,String> dataMap = new HashMap<String,String>();
			
			XSSFRow row = sheet.getRow(i);
			
			for(int j=row.getFirstCellNum(); j<row.getLastCellNum(); j++){
				String key = titleRow.getCell(j).getStringCellValue();
				String value = getCellValue(row.getCell(j));
				dataMap.put(key, value);
			}
			
			data.add(dataMap);
		}
		
		table.setData(data);
	}

	private static String getCellValue(XSSFCell cell) {
		
		if(cell==null)
			return "";
			
		int cellType = cell.getCellType();
		
		switch (cellType) {
		case XSSFCell.CELL_TYPE_BOOLEAN:
			boolean bValue = cell.getBooleanCellValue();
			return String.valueOf(bValue);
		case XSSFCell.CELL_TYPE_NUMERIC:
			double nValue = cell.getNumericCellValue();
			return String.valueOf(nValue);
		case XSSFCell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case XSSFCell.CELL_TYPE_BLANK:
			return "";
		case XSSFCell.CELL_TYPE_ERROR:
			return "" + cell.getErrorCellValue();
		case XSSFCell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
		}
		//엑셀에 있는 각각의 셀 값들을 String형태로 반환한다. 
		return null;
		
	}
	

	
}
