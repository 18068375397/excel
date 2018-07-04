package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

/**
 * 操作Excel表格的功能类
 */
public class MenuUtil {

    /**
     * 读入excel文件，解析后返回
     *
     * @param
     * @throws IOException
     */
    public static List<String[]> readExcel() throws IOException {
        InputStream is = MenuUtil.class.getResourceAsStream("导数据.xlsx");
        // 获得Workbook工作薄对象
        XSSFWorkbook workbook = new XSSFWorkbook(is);
        // 创建返回对象，把每行中的值作为一个数组，所有行作为一个集合返回
        List<String[]> list = new ArrayList<String[]>();
        if (workbook != null) {
            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
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
                for (int rowNum = firstRowNum ; rowNum <= lastRowNum; rowNum++) {
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
                    for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
                        XSSFCell cell = row.getCell(cellNum);
                        cells[cellNum] = getCellValue(cell);
                    }
                    list.add(cells);
                }
            }
            workbook.close();
            is.close();
        }
        return list;
    }

    public static String getCellValue(XSSFCell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }

        CellType ct = cell.getCellTypeEnum();
        if (CellType.STRING == ct) {
            return cell.getStringCellValue();
        } else if (CellType.BLANK == ct) {
            return "";
        } else if (CellType.BOOLEAN == ct) {
            return cell.getBooleanCellValue() + "";
        } else if (CellType.NUMERIC == ct) {
            return cell.getNumericCellValue() + "";
        } else if (CellType._NONE == ct) {
            return "";
        }

        return cell.getStringCellValue();
    }

    public static void main(String[] args) throws IOException, ClassNotFoundException, SQLException {
//		File file = Menu.class.getResourceAsStream("/com/lavasoft/res/a.txt");
//				new File("/系统菜单.xlsx");
        List<String[]> list = readExcel();

//		String url = "jdbc:mysql://61.172.253.134:3306/myface?useUnicode=true&characterEncoding=utf-8";
//		String user = "root";
//		String passwd = "E10ADC3949BA59ABBE56E057F20F883E";

        String url = "jdbc:sqlserver://rm-uf63w75805gs1fhy0o.sqlserver.rds.aliyuncs.com:3433;database=cdc_dev";
        String user = "cdcs_test";
        String passwd = "CM2017-cdcs";
        Connection conn = null;
        Statement stmt = null;

        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
        conn = DriverManager.getConnection(url, user, passwd);
        conn.setAutoCommit(false);
        stmt = conn.createStatement();

//		String deleteSql = "DELETE FROM  sys_menu";
//		stmt.addBatch(deleteSql);

        System.out.println(list.size());

        for (String[] s : list) {
            // 系统ID level 父类ID 名称 类型 连接地址  管理员权限	普通用户权限
            String insertSql = "insert into product_inout_sn(distcode,distName,cuscode,cusName, docType,docno,docdate,docremark,itemCode,qty,uom,autoPush,factoryCode,   ind_no,reportId,report_code,report_name,productprice,totalprice,report_org_id) VALUES ('"
                    + s[0] + "','" + s[1] + "','" + s[3] + "','" + s[4] + "','" + s[6] + "','" + s[7] + "','" + s[8] + "','" + s[9] + "','" + s[11] + "','" + s[12] + "','" + s[13] + "','" + "0" + "','" + "MDMUD0000000" + "','" + "0" + "','" + "0" + "','" + s[0] + "','" + s[1] + "','" + "0" + "','" + "0" + "','" + s[0] + "')";
            stmt.addBatch(insertSql);
        }
            stmt.executeBatch();
            conn.commit();
            stmt.close();
            conn.close();

            System.out.println("---------------导入结束--------------------");

        }


}