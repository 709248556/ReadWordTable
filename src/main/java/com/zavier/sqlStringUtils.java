package com.zavier;

import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * @author yanlianglong
 * @Title: sqlStringUtils.java
 * @Package com.zavier
 * @Description:
 * @date 2020/4/13 17:09
 */
public class sqlStringUtils {
    /**
     * 为string前后添加`,如Fbank_id ---> `Fbank_id`
     */
    public static String addSuffAndPre(String soure) {
        StringBuffer s = new StringBuffer(soure);
        s.append("`");
        s.insert(0, "`");
        return s.toString();
    }
    /**
     * 为string前后添加',如 银行系统编号 ---> '银行系统编号'
     */
    public static String addSuffAndPre2(String soure) {
        StringBuffer s = new StringBuffer(soure);
        s.append("'");
        s.insert(0, "'");
        return s.toString();
    }

    /*
     * 创建 sql第一句 ：CREATE TABLE IF NOT EXISTS`bz_bank`(
     */
    public static String getSql1(String s) {
        StringBuffer sql = new StringBuffer("CREATE TABLE IF NOT EXISTS ");
        sql.append(addSuffAndPre(s.toLowerCase()));
        sql.append("(");
        sql.append("\n");
        return sql.toString();
    }

    /*
     * 字段类型 : int(10) ---> int(10)
     *           varchar(30) ---> varchar(30)
     */
    public static String getDataType(String s) {
        return s;
    }

    /*
     * 是否可为空 : NOT NULL
     */
    public static String isNaN(String s) {
        if (s.equals("N")) return "NOT NULL";
        else return s;
    }

    /*
     * 是否为主键自增
     */
    public static String isKeyAndAutoIncrement(String s) {
        if ("主键自增".equals(s)) return "PRIMARY KEY AUTO_INCREMENT ";
        else return "";
    }

    /*
     * 注释
     */
    public static String COMMENT(String s) {
        StringBuffer result = new StringBuffer();
        result.append("COMMENT ");
        result.append(addSuffAndPre2(s));
        return result.toString();
    }

    public static String getSql2(XWPFTable table){
        int tableRowsSize = table.getRows().size()-1;
        StringBuffer tableStr = new StringBuffer();
        for (int i = 1; i <= tableRowsSize; i++) {
            tableStr.append(addSuffAndPre(table.getRow(i).getCell(1).getText()));//字段名
            tableStr.append(" ");
            tableStr.append(getDataType(table.getRow(i).getCell(3).getText()));//字段类型
            tableStr.append(" ");
            tableStr.append(isNaN(table.getRow(i).getCell(5).getText()));//是否可为空
            tableStr.append(" ");
            tableStr.append(isKeyAndAutoIncrement(table.getRow(i).getCell(4).getText()));//说明
            tableStr.append(COMMENT(table.getRow(i).getCell(2).getText()));//字段注释
            if(i != tableRowsSize) tableStr.append(",");
            tableStr.append("\n");
        }
        tableStr.append(")");
        tableStr.append("ENGINE=InnoDB DEFAULT CHARSET=utf8;");
        return tableStr.toString();
    }

    public static String getIndex(XWPFTable table){
        int tableRowsSize = table.getRows().size()-1;
        StringBuffer tableStr = new StringBuffer();
        for (int i = 1; i <= tableRowsSize; i++) {
            String s = table.getRow(i).getCell(0).getText();
            String s1 = table.getRow(i).getCell(1).getText();
            String s2 = table.getRow(i).getCell(2).getText();
            System.out.println();
        }
        return tableStr.toString();
    }
}
