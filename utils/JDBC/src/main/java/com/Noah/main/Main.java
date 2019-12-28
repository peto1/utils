package com.Noah.main;

import com.Noah.utils.ExcelTools;
import java.io.File;
import java.sql.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws ClassNotFoundException {
        List<Map<String, Object>> list = ExcelTools.readExcel(
                new File("C:\\Users\\taobi\\Desktop\\工作簿(2).xlsx"),
                "xlsx",
                null,
                new String[]{"用户名", "角色"});
        int count = 0;
        List<String> check = new ArrayList<>();
        List<String> names = getName();
        List<String> nonexist = new ArrayList<>();
        Map<String,String> userCode = getUserCode();
        Map<String,String> roleCode = getRoleCode();
        try {
            Class.forName("com.mysql.jdbc.Driver");
            String url="jdbc:mysql://172.20.62.80:3307/auth_server";
            String username="root";
            String userpwd="abc123456";
            Connection conn = DriverManager.getConnection(url,username,userpwd);

            conn.setAutoCommit(false);
            PreparedStatement ps = conn.prepareStatement("insert into auth_contact_user_role (user_code, role_code) values(?,? )");


            for (Map<String, Object> map : list) {
                if (list.contains(map.get("用户名").toString()+ map.get("角色").toString()))
                    continue;
                if (!names.contains(map.get("用户名").toString())){
                    if (!nonexist.contains(map.get("用户名")))
                        nonexist.add(map.get("用户名").toString());
                    continue;
                }
                check.add(map.get("用户名").toString()+ map.get("角色").toString());
                ps.setString(1,userCode.get(map.get("用户名")));
                ps.setString(2,roleCode.get(map.get("角色")));
                ps.executeUpdate();
                System.out.println(count++);
            }
            conn.commit();
            ps.close();
            conn.close();
            for (String s : nonexist) {
                System.out.println(s);
            }
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (SQLException e) {
            e.printStackTrace();
        }

        ExcelTools.writeLocalFile(
                new File("C:\\Users\\taobi\\Desktop\\out.xls"),
                ExcelTools.getHSSFWookbook(
                        list,
                        new String[]{"用户名", "角色", "success"},
                        new String[]{"user_code", "role_code", "success"}
                )
        );
    }

    public static List getName() throws ClassNotFoundException {
        List<String> names = new ArrayList<>();
        Class.forName("com.mysql.jdbc.Driver");
        String url="jdbc:mysql://172.20.62.80:3307/auth_server";
        String username="root";
        String userpwd="abc123456";
        try {
            Connection conn = DriverManager.getConnection(url,username,userpwd);
            Statement statement = conn.createStatement();
            ResultSet rs = statement.executeQuery("select username from auth_user");
            while(rs.next()){
                names.add(rs.getString("username"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return names;
    }

    public static Map getUserCode() throws ClassNotFoundException {
        Map<String,String> userCode = new HashMap<>();
        Class.forName("com.mysql.jdbc.Driver");
        String url="jdbc:mysql://172.20.62.80:3307/auth_server";
        String username="root";
        String userpwd="abc123456";
        try {
            Connection conn = DriverManager.getConnection(url,username,userpwd);
            Statement statement = conn.createStatement();
            ResultSet rs = statement.executeQuery("select code, username from auth_user");
            while(rs.next()){
                userCode.put(rs.getString("username"),rs.getString("code"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return userCode;
    }

    public static Map getRoleCode() throws ClassNotFoundException {
        Map<String,String> roleCode = new HashMap<>();
        Class.forName("com.mysql.jdbc.Driver");
        String url="jdbc:mysql://172.20.62.80:3307/auth_server";
        String username="root";
        String userpwd="abc123456";
        try {
            Connection conn = DriverManager.getConnection(url,username,userpwd);
            Statement statement = conn.createStatement();
            ResultSet rs = statement.executeQuery("select code, name from auth_role");
            while(rs.next()){
                roleCode.put(rs.getString("name"),rs.getString("code"));
            }
        } catch (SQLException e) {
            e.printStackTrace();
        }
        return roleCode;
    }
}
