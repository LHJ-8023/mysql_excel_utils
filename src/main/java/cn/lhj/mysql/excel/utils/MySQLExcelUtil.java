package cn.lhj.mysql.excel.utils;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.*;

/**
 * 将MySQL中的数据导出到Excel表格中
 */
public class MySQLExcelUtil {

    private static final Logger log = LoggerFactory.getLogger(MySQLExcelUtil.class);

    private static final String SELECT_SQL = "SELECT * FROM %s";

    /**
     * 将数据库中的数据导出到excel文件中
     * 导出文件名为: 数据库名_export.xls
     * @param ip 数据库ip地址
     * @param port 数据库端口
     * @param database 数据库名
     * @param username 连接用户名
     * @param password 连接密码
     * @param filePath 导出文件保存路径 -- 不含文件名
     * @param excludeTables 不作导出的数据库表名
     * @throws SQLException
     */
    public static void exportFromMySQLToExcel(String ip, String port, String database,
                                              String username, String password, String filePath, String ...excludeTables) throws SQLException {

        Connection connection = MySQLConnectionUtil.getConnection(ip, port, database, username, password);
        if(connection == null){
            log.error("获取数据库连接失败");
            throw new RuntimeException("数据库连接失败");
        }

        if(filePath == null){
            log.error("文件导出路径不能为空");
            throw new RuntimeException("文件导出路径不能为空");
        }

        ArrayList<String> excludes = new ArrayList<>(Arrays.asList(excludeTables));

        //获取数据库所有表信息
        DatabaseMetaData metaData = connection.getMetaData();
        ResultSet rs = metaData.getTables(null, null, null, new String[]{"TABLE"});

        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();

        while(rs.next()){
            String tableName = rs.getString("TABLE_NAME");
            if(!excludes.contains(tableName)){
                //获取表中的数据
                Map<Integer, Map<String, Object>> datas = getAllDataFromTable(tableName, connection);
                if(datas == null) break;

                //获取字段名和类型
                Map<String, String> nameType = getDataNameAndType(tableName, connection);
                if(nameType == null) break;

                //创建一张表
                HSSFSheet sheet = workbook.createSheet(tableName);
                //初始化第一行 显示表中的列
                Object[] columnNames = nameType.keySet().toArray();
                HSSFRow firstRow = sheet.createRow(0);
                for(int i = 0; i < columnNames.length; i++){
                    firstRow.createCell(i + 1).setCellValue((String)columnNames[i]);
                }

                //向工作簿写入数据
                for(int row = 0; row < datas.size(); row++){
                    Map<String, Object> data = datas.get(row);
                    HSSFRow rows = sheet.createRow(row + 1);
                    //设置第一列数据为数据行数
                    rows.createCell(0).setCellValue(row + 1);
                    for(int col = 0; col < data.size(); col++){
                        Object value = data.get((String) columnNames[col]);
                        if(value == null) break;
                        rows.createCell(col + 1).setCellValue(value.toString());
                    }
                }

            }
        }
		
		//关闭连接
		MySQLConnectionUtil.close(statement);

        //保存工作簿
        try {
            File file = new File(filePath);
            if(!file.exists()) file.mkdirs();
            filePath = filePath + "\\" + database + "_export.xls";
            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            log.error(e.getLocalizedMessage());
        } catch (IOException e) {
            log.error(e.getMessage());
            e.printStackTrace();
        }

    }

    private static Map<Integer, Map<String, Object>> getAllDataFromTable(String tableName, Connection conn){
        LinkedHashMap<Integer, Map<String, Object>> res = null;
        PreparedStatement statement = null;

        try {
            res = new LinkedHashMap<>();

            //根据表明查询表中所有数据
            statement = conn.prepareStatement(String.format(SELECT_SQL, tableName));
            ResultSet rs = statement.executeQuery();

            //获取表中字段总数
            ResultSetMetaData metaData = rs.getMetaData();
            int columnCount = metaData.getColumnCount();
            int count = 0;

            while(rs.next()){
                LinkedHashMap<String, Object> tmp = new LinkedHashMap<>();
                for(int i = 1; i <= columnCount; i++){
                    String columnName = metaData.getColumnName(i);
                    Object obj = rs.getObject(i);
                    tmp.put(columnName, obj);
                }
                res.put(count++, tmp);
            }
        } catch (SQLException e) {
            e.printStackTrace();
            log.error("查询数据库表: " + tableName + " 信息失败");
            log.error(e.getMessage());
            return null;
        } finally {
            MySQLConnectionUtil.close(statement);
        }

        return res;
    }

    private static Map<String, String> getDataNameAndType(String tableName, Connection conn){
        LinkedHashMap<String, String> res = null;
        PreparedStatement statement = null;

        try {
            statement = conn.prepareStatement(String.format(SELECT_SQL, tableName));
            ResultSet rs = statement.executeQuery();

            res = new LinkedHashMap<>();

            ResultSetMetaData metaData = rs.getMetaData();
            int count = metaData.getColumnCount();

            for(int i = 1; i <= count; i++){
                res.put(metaData.getColumnName(i), metaData.getColumnClassName(i));
            }

        } catch (SQLException e) {
            e.printStackTrace();
            log.error("查询数据库表: " + tableName + " 信息失败");
            log.error(e.getMessage());
            return null;
        } finally {
            MySQLConnectionUtil.close(statement);
        }

        return res;
    }

}
