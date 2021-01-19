package com.zavier;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.List;

/**
 * @author yanlianglong
 * @Title: sqlStringUtils.java
 * @Package com.zavier
 * @Description:
 * @date 2020/4/13 17:09
 */
public class sqlStringUtils {

    /*
     * 获取表名
     */
    private static String getTableName(List<XWPFRun> runs) {
        String s1 = "";
        for (int i = 1; i < runs.size(); i++) {
            s1 += runs.get(i);
        }
        return s1.toLowerCase();
    }

    /**
     * 为string前后添加`,如Fbank_id ---> `Fbank_id`
     *
     * 全部改成小写
     */
    public static String addSuffAndPre(String soure) {
        StringBuffer s = new StringBuffer(soure.toLowerCase());
//        s.append("`");
//        s.insert(0,"`");
        return s.toString();
    }

    /**
     * 为string前后添加',如 银行系统编号 ---> '银行系统编号'
     */
    public static String addSuffAndPre2(String soure) {
        StringBuffer s = new StringBuffer(soure);
        s.append("'");
        s.insert(0,"'");
        return s.toString();
    }

    /*
     * 删除已存在的表
     */
    public static String deleteTable(List<XWPFRun> s) {
        String pre = "if exists(select * from sysobjects where name='TABLE_NAME')\nbegin\ndrop table TABLE_NAME\nend\n";
        pre = pre.replaceAll("TABLE_NAME",getTableName(s));
        return pre;
    }

    /*
     * 创建 sql第一句 ：CREATE TABLE IF NOT EXISTS`bz_bank`(
     */
    public static String getSql1(List<XWPFRun> s) {
        StringBuffer sql = new StringBuffer("CREATE TABLE ");
        sql.append(addSuffAndPre(getTableName(s)));
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
        else if ("主键".equals(s) ) return "";
        else return s;
    }

    /*
     * 是否为主键自增
     */
    public static String isKeyAndAutoIncrement(String s) {
        if ("主键".equals(s)) return "identity(1,1) PRIMARY KEY ";
        else return "";
    }

    /*
     * 注释
     */
    public static String COMMENT(String s) {
        StringBuffer result = new StringBuffer();
//        result.append("COMMENT ");
        result.append(addSuffAndPre2(s));
        return result.toString();
    }

    public static String getSql2(XWPFTable table) {
        int tableRowsSize = table.getRows().size() - 1;
        StringBuffer tableStr = new StringBuffer();
        for (int i = 1; i <= tableRowsSize; i++) {
            tableStr.append(addSuffAndPre(table.getRow(i).getCell(1).getText()));//字段名
            tableStr.append(" ");
            tableStr.append(getDataType(table.getRow(i).getCell(3).getText()));//字段类型
            tableStr.append(" ");
            tableStr.append(isNaN(table.getRow(i).getCell(2).getText()));//是否可为空
            tableStr.append(" ");
            tableStr.append(isKeyAndAutoIncrement(table.getRow(i).getCell(2).getText()));//说明
//            tableStr.append(COMMENT(table.getRow(i).getCell(5).getText()));//字段注释
            if (i != tableRowsSize) tableStr.append(",");
            tableStr.append("\n");
        }
        tableStr.append(")");
//        tableStr.append("ENGINE=InnoDB DEFAULT CHARSET=utf8;");
        return tableStr.toString();
    }

    public static String getIndex(XWPFTable table,String tableName) {
        int tableRowsSize = table.getRows().size() - 1;
        StringBuffer tableStr = new StringBuffer();

        for (int i = 1; i <= tableRowsSize; i++) {
            tableStr.append("ALTER TABLE " + tableName + " ADD ");
            String s = table.getRow(i).getCell(0).getText();
            String s1 = table.getRow(i).getCell(1).getText();
            String s2 = table.getRow(i).getCell(2).getText();
            tableStr.append(s1);
            tableStr.append(" " + addSuffAndPre(s) + " ");
            tableStr.append("(" + addSuffAndPre(s2) + ")");
            tableStr.append(";");
        }
        return tableStr.toString();
    }

    /**
     *获取字段说明
     * @param document
     */
    public static void getColumnSql(XWPFDocument document) {
        String sql = "EXEC sp_addextendedproperty 'MS_Description', N'DESCRIPTION', 'SCHEMA', N'dbo', 'TABLE', N'TABLE_NAME', 'COLUMN', N'FIELD'";
        String tableName = getTableName(document.getParagraphs().get(0).getRuns());
        sql = sql.replace("TABLE_NAME",tableName);
        XWPFTable table = document.getTables().get(0);
        int tableRowsSize = table.getRows().size() - 1;
        for (int i = 1; i <= tableRowsSize; i++) {
            String result = new String(sql);
            String field = addSuffAndPre(table.getRow(i).getCell(1).getText());//字段名
            String comment = table.getRow(i).getCell(5).getText();//字段注释
            result = result.replace("DESCRIPTION",comment);
            result = result.replace("FIELD",field);
            System.out.println(result);
        }
    }

}
