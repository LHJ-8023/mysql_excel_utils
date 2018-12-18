package cn.lhj.mysql.excel.test;

import cn.lhj.mysql.excel.utils.MySQLExcelUtil;
import org.junit.Test;

import java.sql.SQLException;

public class MyTest {

    @Test
    public void fun_1() throws SQLException {
        String ip = "127.0.0.1";
        String port = "3306";
        String username = "root";
        String password = "root";
        String database = "mysql_excel";
        String filePath = "E:\\Git-Project\\Test\\mysql-excel";
        MySQLExcelUtil.exportFromMySQLToExcel(ip, port, database, username, password, filePath);
    }

    @Test
    public void d() throws SQLException {
        //MySQLExcelUtil.exportFromMySQLToExcel(null, null);

        String ip = "127.0.0.1";
        String port = "3306";
        String username = "root";
        String password = "root";
        String database = "mysql_excel";
        String filePath = "E:\\Git-Project\\Test\\mysql-excel\\mysql_excel_export.xls";

        MySQLExcelUtil.importFromExcelToMySQL(ip, port, database, username, password, filePath);


    }

}
