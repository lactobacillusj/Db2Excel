package com.lactobacillusj.main;

import java.sql.Connection;
import java.util.List;

import com.lactovacillusj.util.Tools_upgrade;
import com.lactobacillusj.vo.TableVO;

public class MainApp_E2D {
	public static void main(String[] args) throws Exception {

		String fileName = "C:\\Users\\TA°ø¿ë\\Desktop\\storege_mk\\example\\scott.xlsx";
		
		List<TableVO> tableVOList = Tools_upgrade.excelToVOList(fileName);
		
		
		String driver = "oracle.jdbc.driver.OracleDriver";
		String url = "jdbc:oracle:thin:@localhost:1521:iip";
		String id = "scott";
		String pwd = "tiger";
		
		Connection conn = Tools_upgrade.connectDB(driver, url, id, pwd);

		for(int i=0; i<tableVOList.size(); i++){
			Tools_upgrade.createTable(conn, tableVOList.get(i));
			Tools_upgrade.insertData(conn, tableVOList.get(i));
		}
	}
}
