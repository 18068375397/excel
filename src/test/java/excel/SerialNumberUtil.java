package excel;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

public class SerialNumberUtil {
    /**
     * 解析excel，返回序列号列表
     *
     * @param
     * @throws IOException
     */
    public static List<String> readExcel() throws IOException {
        InputStream is = MenuUtil.class.getResourceAsStream("六月召回产品(1).xlsx");
        // 获得Workbook工作薄对象
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        // 创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
        List<String> list = new ArrayList<String>();
        StringBuffer sb = new StringBuffer("");
        if (workbook != null) {
//            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            for (int sheetNum = 0; sheetNum < 1; sheetNum++) {
                // workbook.getSheet("菜单");
                // 获得当前sheet工作表
                XSSFSheet sheet = workbook.getSheetAt(sheetNum);
                if (sheet == null) {
                    continue;
                }
                // 获得当前sheet的开始行
                int firstRowNum = sheet.getFirstRowNum();
                // 获得当前sheet的结束行
                int lastRowNum = sheet.getLastRowNum();
                // 循环除了第一行的所有行
                for (int rowNum = 2; rowNum <= lastRowNum; rowNum++) {
                    // 获得当前行
                    XSSFRow row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    // 获得当前行的开始列
                    int firstCellNum = row.getFirstCellNum();
                    // 获得当前行的列数
                    int lastCellNum = row.getLastCellNum();
                    String[] cells = new String[row.getLastCellNum()];
                    // 循环当前行
//                    for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
//                        XSSFCell cell = row.getCell(cellNum);
//                        cells[cellNum] = MenuUtil.getCellValue(cellNum, cell);
//                    }

                    list.add(MenuUtil.getCellValue(row.getCell(10)));
                }
            }
            workbook.close();
            is.close();
        }
        System.out.println(sb.toString());
        return list;
    }

    /**
     * 读取excel数据，根据serialNumber批量更新数据
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        List<String> list = readExcel();

        String url = "jdbc:sqlserver://rm-uf63w75805gs1fhy0o.sqlserver.rds.aliyuncs.com:3433;database=cdcs";
        String user = "yzcdcs";
        String passwd = "YZcdcs-2017";
        Connection conn = null;
        Statement stmt = null;

        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
        conn = DriverManager.getConnection(url, user, passwd);
        conn.setAutoCommit(false);
        stmt = conn.createStatement();


        System.out.println(list.size());

        for (String s : list) {
            if (!s.isEmpty()) {
                // 系统ID level 父类ID 名称 类型 连接地址  管理员权限	普通用户权限
                //            String insertSql = "insert into product_inout_sn(distcode,distName,cuscode,cusName, docType,docno,docdate,docremark,itemCode,qty,uom,autoPush,factoryCode,   ind_no,reportId,report_code,report_name,productprice,totalprice,report_org_id) VALUES ('"
                //                    + s[0] + "','" + s[1] + "','" + s[3] + "','" + s[4] + "','" + s[6] + "','" + s[7] + "','" + s[8] + "','" + s[9] + "','" + s[11] + "','" + s[12] + "','" + s[13] + "','" + "0" + "','" + "MDMUD0000000" + "','" + "0" + "','" + "0" + "','" + s[0] + "','" + s[1] + "','" + "0" + "','" + "0" + "','" + s[0] + "')";
                String updateSql = "UPDATE product_sn SET status = '预入库',cuscode = 'MDMUD025535',cusName = '上海九州通医疗器械供应链有限公司' WHERE serialnumber = '" + s + "';";
                System.out.println(updateSql);
                stmt.addBatch(updateSql);
            }
        }
        stmt.executeBatch();
        conn.commit();
        stmt.close();
        conn.close();

        System.out.println("---------------修改结束--------------------");

    }
}
