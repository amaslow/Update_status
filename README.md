# Update_status
Export from MySQL database to Excel file

Utils.java:

package update_status;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class Utils {

    private static Connection con;
    private static String Driver = "com.mysql.jdbc.Driver";
    private static String url = "jdbc:mysql://{ip address}:3306/{database name}";
    private static String username = "{username}";
    private static String password = "{password}";

    public static Connection getConnection() {
        try {
            Class.forName(Driver);
            con = DriverManager.getConnection(url, username, password);
        } catch (ClassNotFoundException e) {
            // TODO Auto - generated catch block
            e.printStackTrace();
        } catch (SQLException e) {
            // TODO Auto-generated catch block 
            e.printStackTrace();
        }
        return con;
    }

    public static void closeDB(ResultSet rs, Statement st, Connection con) {
        try {
            if (rs != null) {
                rs.close();
                rs = null;
            }
            if (st != null) {
                st.close();
                st = null;
            }
            if (con != null) {
                con.close();
                con = null;
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }

    }
}
