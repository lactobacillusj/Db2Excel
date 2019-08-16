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

<<<<<<< HEAD
		Connection conn = tools.connectDB(driver, url, id, pw); // Connection
		List<String> tableList = tools.getTableList(conn);
		XSSFWorkbook wb = new XSSFWorkbook();
=======
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
>>>>>>> refs/remotes/origin/master

<<<<<<< HEAD
		for (int count = 0; count < tableList.size(); count++) {
			String tableName = tableList.get(count);
			List<?> tableInfo = tools.db2List(conn, tableName);
			tools.putTableInfo(wb, tableName, tableInfo);
=======
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
>>>>>>> refs/remotes/origin/master
		}
////
		String fileName = "C:\\Users\\TA����\\Desktop\\storege_mk\\example\\db2excel.xls";
		tools.writeExcel(wb, fileName);
	}
}