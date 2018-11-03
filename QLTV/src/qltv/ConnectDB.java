/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package qltv;
import java.sql.*;
/**
 *
 * @author Pham Ngoc Minh
 */

public class ConnectDB {
	public Connection getConnect() {
		Connection connection = null;
		String url = "jdbc:mysql://localhost:3306/qltv?autoReconnect=true&useSSL=false";
		String user = "root";
		String passwd = "1234";
		try {
			
			connection = DriverManager.getConnection(url, user, passwd);
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
		return connection;
	}

}
