package firstDemo;


import java.io.InputStream;
import java.sql.*;
import java.util.Properties;

public class JdbcUtils {
    private static String driver=null;
    private static String url=null;
    private static String username=null;
    private static String password=null;

    static {
        try{
            InputStream in = firstDemo.JdbcUtils.class.getClassLoader().getResourceAsStream("db.properties");
            Properties pr=new Properties();
            pr.load(in);

            driver=pr.getProperty("driver");
            url=pr.getProperty("url");
            username=pr.getProperty("username");
            password=pr.getProperty("password");

            //1、加载驱动
            Class.forName(driver);

        }catch (Exception e) {
            e.printStackTrace();
        }
    }

    //2、获取连接
    public static Connection getConnection() throws SQLException {
        return DriverManager.getConnection(url, username, password);
    }

    //6、释放资源
    public static void release(Connection conn, Statement st, ResultSet rs){
        if(rs!=null){
            try {
                rs.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if(st!=null){
            try {
                st.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
        if(conn!=null){
            try {
                conn.close();
            } catch (SQLException e) {
                e.printStackTrace();
            }
        }
    }

}
