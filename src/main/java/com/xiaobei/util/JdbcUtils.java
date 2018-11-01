package com.xiaobei.util;

import org.apache.commons.dbcp.BasicDataSourceFactory;

import javax.sql.DataSource;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;
/**
 * 基于mysql的jdbc工具类
 * @author Administrator
 *
 */
public final class JdbcUtils {//使用开源的数据源：dbcp：databaseConnectionPool：数据库连接池
    /**
     * 导入的三个包介绍：
     * 			commons-collections-3.1.jar：里面放的是apache自己的集合类，我们自己创建的数据源使用的是java的集合：LinkedList
     * 			commons-dbcp-1.2.2.jar：针对数据库连接
     * 			commons-pool.jar：通用的连接池：往commons-collections-3.1.jar中的集合类存放connection，这个pool是通用的，也可以放其他的资源
     */
    private static DataSource dataSource;
    private JdbcUtils() {
    }

    static {
        try {
            Class.forName("com.mysql.jdbc.Driver");
            Properties pro = new Properties();
            InputStream in = JdbcUtils.class.getClassLoader().getResourceAsStream("dbcpconfig.properties");
            pro.load(in);
            dataSource = BasicDataSourceFactory.createDataSource(pro );
        } catch (Exception e) {
            throw new ExceptionInInitializerError(e);
        }
    }

    public static Connection getConnection() throws SQLException {
        return dataSource.getConnection();//从连接池中获取已经创建好的连接
    }

    public static void free(ResultSet rs, Statement st, Connection conn) {
        try {
            if (rs != null)
                rs.close();
        } catch (SQLException e) {
            e.printStackTrace();
        } finally {
            try {
                if (st != null)
                    st.close();
            } catch (SQLException e) {
                e.printStackTrace();
            } finally {
                if (conn != null)
                    try {
                        conn.close();/**关闭的时候 ，其实是将该连接放回连接池中*/
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
            }
        }
    }
    /**测试创建的连接
     * @throws SQLException */
    public static void main(String[] args) throws SQLException {
        for(int i = 0;i<16;i++){
            Connection conn = JdbcUtils.getConnection();
            System.out.println(conn.hashCode());
//            if (i >= 9) {
//                JdbcUtil.free(null, null, conn);
//            }
        }
    }
}

