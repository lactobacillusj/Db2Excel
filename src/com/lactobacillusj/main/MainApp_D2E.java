package com.lactobacillusj.main;

import java.sql.Connection;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.lactovacillusj.util.Tools_upgrade;
import com.lactobacillusj.vo.TableVO;

public class MainApp_D2E {

	public static void main(String[] args) throws Exception {

		String driver = "oracle.jdbc.driver.OracleDriver";
		String url = "jdbc:oracle:thin:@localhost:1521:iip";
		String id = "scott";
		String pwd = "tiger";
		
		Connection conn = Tools_upgrade.connectDB(driver, url, id, pwd);
		
		//========== 2. ���̺� ��� �������� =============================

		String sql = "SELECT TABLE_NAME FROM ALL_TABLES WHERE OWNER = 'SCOTT'";
		
		List<String> tableNameList = Tools_upgrade.getListFirstColumn(conn, sql);
		
		//========== 3. ��ũ�� ����� ==========================
		XSSFWorkbook wb = new XSSFWorkbook();
		
		XSSFSheet tableInfoSheet = Tools_upgrade.setTableInfoSheet(wb);
		int tableInfoRowCnt = 1;
		
		for(int i = 0 ; i < tableNameList.size() ; i++){

			//========== 3-1. ���̺� ���� �������� =============================
			TableVO table = new TableVO(tableNameList.get(i));
			
			table.loadColumn(conn);
			table.loadData(conn);
			
			//========== 3-2. ��ũ�Ͽ� ��� =============================
			
			Tools_upgrade.addVOToSheet(wb, table);

			//========== 3-3. ���̺����� ��Ʈ�� ��� ============================
			tableInfoRowCnt = Tools_upgrade.addTableInfoRow(tableInfoSheet, tableInfoRowCnt, table);
			
		}

		//========== 4. �������Ϸ� ����� =============================
		String fileName = "C:\\Users\\TA����\\Desktop\\storege_mk\\example\\scott.xlsx";
		Tools_upgrade.writeExcel(wb, fileName);
	}
}