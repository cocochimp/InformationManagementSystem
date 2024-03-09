package JDBCTest;

import java.sql.*;

public class JDBCText01 {
    public static void main(String[] args) throws ClassNotFoundException, SQLException {
        //1、加载驱动
        //DriverManager.registerDriver(new com.mysql.jdbc.Driver());
        Class.forName("com.mysql.jdbc.Driver"); //固定写法，加载驱动

        //2、获取连接(用户信息和url)
        String url="jdbc:mysql://localhost:3306/test?" +
                "useUnicode=true&characterEncoding=utf8&useSSL=true";
        String username="root";
        String password="123456";

        //3、获取数据库对象
        Connection connection= DriverManager.getConnection(url,username,password);

        //4、执行SQl语句
        Statement statement = connection.createStatement();

        //5、处理查询结果集
        String sql="select * from emp";
        ResultSet resultSet=statement.executeQuery(sql);

        while(resultSet.next()){
            System.out.println("ename:"+resultSet.getObject("ename"));
        }

        //6、释放资源
        resultSet.close();
        statement.close();
        connection.close();
    }
}
