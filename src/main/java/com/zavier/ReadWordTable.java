package com.zavier;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

/**
 * 读取word中的表格，包括复杂表格（合并的单元格）
 */
public class ReadWordTable {

    /**
     * 保存生成HTML时需要被忽略的单元格
     */
    private List<String> omitCellsList = new ArrayList<>();

    /**
     * 生成忽略的单元格列表中的格式
     *
     * @param row
     * @param col
     * @return
     */
    public String generateOmitCellStr(int row, int col) {
        return row + ":" + col;
    }

    /**
     * 获取当前单元格的colspan（列合并）的列数
     *
     * @param tcPr 单元格属性
     * @return
     */
    public int getColspan(CTTcPr tcPr) {
        // 判断是否存在列合并
        CTDecimalNumber gridSpan = null;
        if ((gridSpan = tcPr.getGridSpan()) != null) { // 合并的起始列
            // 获取合并的列数
            BigInteger num = gridSpan.getVal();
            return num.intValue();
        } else { // 其他被合并的列或正常列
            return 1;
        }
    }

    /**
     * 获取当前单元格的rowspan（行合并）的行数
     *
     * @param table 表格
     * @param row   行值
     * @param col   列值
     * @return
     */
    public int getRowspan(XWPFTable table, int row, int col) {

        XWPFTableCell cell = table.getRow(row).getCell(col);
        // 正常独立单元格
        if (!isContinueRow(cell) && !isRestartRow(cell)) {
            return 1;
        }
        // 当前单元格的宽度
        int cellWidth = getCellWidth(table, row, col);
        // 当前单元格距离左侧边框的距离
        int leftWidth = getLeftWidth(table, row, col);

        // 用户保存当前单元格行合并的单元格数-1（因为不包含自身）
        List<Boolean> list = new ArrayList<>();
        getRowspan(table, row, cellWidth, leftWidth, list);

        return list.size() + 1;
    }

    private void getRowspan(XWPFTable table, int row, int cellWidth, int leftWidth,
                            List<Boolean> list) {
        // 已达到最后一行
        if (row + 1 >= table.getNumberOfRows()) {
            return;
        }
        row = row + 1;
        int colsNum = table.getRow(row).getTableCells().size();
        // 因为列合并单元格可能导致行合并的单元格并不在同一列，所以从头遍历列，通过属性、宽度以及距离左边框间距来判断是否是行合并
        for (int i = 0; i < colsNum; i++) {
            XWPFTableCell testTable = table.getRow(row).getCell(i);
            // 是否为合并单元格的中间行（包括结尾行）
            if (isContinueRow(testTable)) {
                // 是被上一行单元格合并的单元格
                if (getCellWidth(table, row, i) == cellWidth
                        && getLeftWidth(table, row, i) == leftWidth) {
                    list.add(true);
                    // 被合并的单元格在生成html时需要忽略
                    addOmitCell(row, i);
                    // 去下一行继续查找
                    getRowspan(table, row, cellWidth, leftWidth, list);
                    break;
                }
            }
        }
    }

    /**
     * 判断是否是合并行的起始行单元格
     *
     * @param tableCell
     * @return
     */
    public boolean isRestartRow(XWPFTableCell tableCell) {
        CTTcPr tcPr = tableCell.getCTTc().getTcPr();
        if (tcPr.getVMerge() == null) {
            return false;
        }
        if (tcPr.getVMerge().getVal() == null) {
            return false;
        }
        if (tcPr.getVMerge().getVal().toString().equalsIgnoreCase("restart")) {
            return true;
        }
        return false;
    }

    /**
     * 判断是否是合并行的中间行单元格（包括结尾的最后一行的单元格）
     *
     * @param tableCell
     * @return
     */
    public boolean isContinueRow(XWPFTableCell tableCell) {
        CTTcPr tcPr = tableCell.getCTTc().getTcPr();
        if (tcPr.getVMerge() == null) {
            return false;
        }
        if (tcPr.getVMerge().getVal() == null) {
            return true;
        }
        return false;
    }

    public int getLeftWidth(XWPFTable table, int row, int col) {
        int leftWidth = 0;
        for (int i = 0; i < col; i++) {
            leftWidth += getCellWidth(table, row, i);
        }
        return leftWidth;
    }

    public int getCellWidth(XWPFTable table, int row, int col) {
        BigInteger width = table.getRow(row).getCell(col).getCTTc().getTcPr().getTcW().getW();
        return width.intValue();
    }

    /**
     * 添加忽略的单元格(被行合并的单元格，生成HTML时需要忽略)
     *
     * @param row
     * @param col
     */
    public void addOmitCell(int row, int col) {
        String omitCellStr = generateOmitCellStr(row, col);
        omitCellsList.add(omitCellStr);
    }



    /**
     * 转化数据库可识别的类型，如int(10) ----> int;varchar(30) ----> varchar(30)
     */
    public static String type(String soure) {
        StringBuffer s = new StringBuffer(soure);
        s.append("`");
        s.insert(0, "`");
        return s.toString();
    }

    public boolean isOmitCell(int row, int col) {
        String cellStr = generateOmitCellStr(row, col);
        return omitCellsList.contains(cellStr);
    }

    public StringBuffer getTable(XWPFTable table){
        int tableRowsSize = table.getRows().size()-1;
        StringBuffer tableToHtmlStr = new StringBuffer();
        for (int i = 1; i <= tableRowsSize; i++) {
            tableToHtmlStr.append(sqlStringUtils.addSuffAndPre(table.getRow(i).getCell(1).getText()));//字段名
            tableToHtmlStr.append(" ");
            tableToHtmlStr.append(sqlStringUtils.getDataType(table.getRow(i).getCell(3).getText()));//字段类型
            tableToHtmlStr.append(" ");
            tableToHtmlStr.append(sqlStringUtils.isNaN(table.getRow(i).getCell(5).getText()));//是否可为空
            tableToHtmlStr.append(" ");
            tableToHtmlStr.append(sqlStringUtils.isKeyAndAutoIncrement(table.getRow(i).getCell(4).getText()));//说明
            tableToHtmlStr.append(sqlStringUtils.COMMENT(table.getRow(i).getCell(2).getText()));//字段注释
            if(i != tableRowsSize) tableToHtmlStr.append(",");
            tableToHtmlStr.append("\n");
        }
        return tableToHtmlStr;
    }

    public void clearTableInfo() {
        // System.out.println(omitCellsList);
        omitCellsList.clear();
    }

    public static void main1(String[] args) {
        ReadWordTable readWordTable = new ReadWordTable();

        try (FileInputStream fileInputStream = new FileInputStream("E:\\java\\ReadWordTable-master\\src\\main\\resources\\bank.docx");
             XWPFDocument document = new XWPFDocument(fileInputStream);) {

            //方法2
            XWPFWordExtractor extractor = new XWPFWordExtractor(document);
            String text = extractor.getText();
            System.out.println(text);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main2(String[] args) {
        try (FileInputStream fileInputStream = new FileInputStream("E:\\java\\ReadWordTable-master\\src\\main\\resources\\bank.docx");
             XWPFDocument document = new XWPFDocument(fileInputStream);) {
            String sql1 = sqlStringUtils.getSql1(document.getParagraphs().get(0).getRuns().get(1).toString());
            List<XWPFTable> tables = document.getTables();
            String sql2 = sqlStringUtils.getSql2(tables.get(0));
            System.out.println(sql1+sql2);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static void main(String[] args) {
        try (FileInputStream fileInputStream = new FileInputStream("E:\\java\\ReadWordTable-master\\src\\main\\resources\\bank.docx");
             XWPFDocument document = new XWPFDocument(fileInputStream);) {
            List<XWPFTable> tables = document.getTables();
            String index = sqlStringUtils.getIndex(tables.get(1),document.getParagraphs().get(0).getRuns().get(1).toString());
            System.out.println(index);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
