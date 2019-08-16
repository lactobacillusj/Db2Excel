package com.lactobacillusj.main;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class MainApp_EP {

	public static void main(String[] args) {

		Connection conn; // Connection
		ResultSet rs;
		Statement stmt = null;
		ResultSetMetaData metaData = null;

		try {
			String dbURL = "jdbc:oracle:thin:@localhost:1521:IIP";
			String dbID = "scott";
			String dbPassword = "tiger";
			Class.forName("oracle.jdbc.driver.OracleDriver"); // 디비 접속을 도와주는 mysql접속 드라이버
			conn = DriverManager.getConnection(dbURL, dbID, dbPassword); // dbURL,dbID, dbPassword를 사용해서 접속한다. conn에 접속
																			// 정보가 담김
			stmt = conn.createStatement();
			String sql = "SELECT * FROM tabs";
			rs = stmt.executeQuery(sql);
			metaData = rs.getMetaData();

			// 각 행을 읽어 리스트에 저장한다.
			int sizeOfcolumn = metaData.getColumnCount();
			String column;
			List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
			Map<String, Object> map;
			//
			while (rs.next()) {
				map = new HashMap<String, Object>();

				for (int indexOfcolumn = 0; indexOfcolumn < sizeOfcolumn; indexOfcolumn++) {
					column = metaData.getColumnName(indexOfcolumn + 1);
					map.put(column, rs.getString(column));
				}
				list.add(map);
			}
			@SuppressWarnings("resource")
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("DEPT");
			HSSFRow row = sheet.createRow(1);
			HSSFRow titleRow = sheet.createRow(0);
			HSSFCell cell;
			HSSFCell celltt;
			int haderIdx = 0;
			
			int cell_create = 0;
			int row_num = 1;

			for (Map<String, Object> map1 : list) {
				System.out.println("=========================================");
				Iterator<String> it = map1.keySet().iterator();

				while (it.hasNext()) {
					String value = (String) map1.get(metaData.getColumnName(haderIdx % sizeOfcolumn + 1));
					String columnN = metaData.getColumnName(haderIdx % sizeOfcolumn + 1);
					System.out.println("columnN : " + columnN);

					haderIdx++;

					if (haderIdx < sizeOfcolumn+1) {
						celltt = titleRow.createCell((cell_create) % sizeOfcolumn);
						celltt.setCellValue(columnN);
						cell = row.createCell((cell_create++) % sizeOfcolumn);
						cell.setCellValue(value);

					} else {
						cell = row.createCell((cell_create++) % sizeOfcolumn);
						cell.setCellValue(value);
					}
					
				}
				row = sheet.createRow(++row_num);

				// System.out.println("=========================================");
			}
			File file = new File("C:\\Users\\TA공용\\Desktop\\storege_mk\\example\\examplewrite.xls");
			FileOutputStream fileOut = new FileOutputStream(file);
			workbook.write(fileOut);
			System.out.println("파일 출력 완료");
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
