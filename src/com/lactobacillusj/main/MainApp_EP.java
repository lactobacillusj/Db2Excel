package com.lactobacillusj.main;

import java.sql.Connection;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lactovacillusj.util.Tools;

public class MainApp_EP {

	public static void main(String[] args) throws Exception {

		Tools tools = new Tools();
		String driver = "oracle.jdbc.driver.OracleDriver";
		String url = "jdbc:oracle:thin:@localhost:1521:IIP";
		String id = "scott";
		String pw = "tiger";


		Connection conn = tools.connectDB(driver, url, id, pw); // Connection
		List<String> tableList = tools.getTableList(conn);
		XSSFWorkbook wb = new XSSFWorkbook();
		for(int i = 0 ; i < tableList.size();i++) {
			String tableName = tableList.get(i);
			List<?> tableInfo = tools.db2List(conn, tableName);
			tools.putTableInfo(wb, tableName, tableInfo);
		}
		String fileName = "C:\\Users\\TA°ø¿ë\\Desktop\\storege_mk\\example\\db2excel.xlsx";
		tools.writeExcel(wb, fileName);
	}
}