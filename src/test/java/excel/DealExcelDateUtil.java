package excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class DealExcelDateUtil {

    //当前文件已经存在
    private static String excelPath = "C:/Users/asus/Desktop/deal.xlsx";
    // 起始序列号所在单元格(从0开始)
    private static int BCellNum = 5;
    private static int ECellNum = 6;
    // 根据excel中数据取（未做智能截取）
    private static String per = "DJTR00000";

    public static void main(String[] args) throws Exception {
        dealExcel();
//        System.out.println(per.split("000").toString());
    }

    /**
     * 处理补打excel（指定excel中序列号行对比）
     *
     * @param
     * @throws IOException
     */
    public static void dealExcel() throws IOException {
        File file = new File(excelPath);
        InputStream is = new FileInputStream(file);

        if (is != null) {
            // 获得Workbook工作薄对象
            XSSFWorkbook workbook = new XSSFWorkbook(is);
            if (workbook != null) {
                //            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                // 只处理第一个sheet
                for (int sheetNum = 0; sheetNum < 1; sheetNum++) {
                    // workbook.getSheet("菜单");
                    // 获得当前sheet工作表
                    XSSFSheet sheet = workbook.getSheetAt(sheetNum);
                    if (sheet == null) {
                        continue;
                    }
                    // 获得当前sheet的开始行
                    int firstRowNum = sheet.getFirstRowNum();
                    //                // 获得当前sheet的结束行
                    //                int lastRowNum = sheet.getLastRowNum();
                    // 循环除了第一行的所有行
                    for (int rowNum = 1; ; rowNum++) {
                        // 获得当前行
                        XSSFRow row = sheet.getRow(rowNum);
                        if (row == null) {
                            break;
                        }

                        XSSFCell cell1 = row.getCell(BCellNum);
                        XSSFCell cell2 = row.getCell(ECellNum);
                        // 获取起始序列号
                        String beginNumber = MenuUtil.getCellValue(cell1);
                        // 获取结束序列号
                        String endNumber = MenuUtil.getCellValue(cell2);
                        // 获取数量
                        //                    String qty = "";

                        int begin = 0;
                        int end = 0;

                        // 序列号不同，需要处理数据并插入
                        if (!beginNumber.equals(endNumber)) //   && !qty.equals("1")
                        {
                            begin = Integer.valueOf(beginNumber.substring(beginNumber.length() - 6));
                            end = Integer.valueOf(endNumber.substring(endNumber.length() - 6));
                            cell2.setCellValue(beginNumber);
                            System.out.println("第" + (rowNum + 1) + "行数据处理，共" + (end - begin) + "条。起始序列号为：" + beginNumber + "，结束序列号为：" + endNumber);

                            // 复制所需插入数据
                            XSSFRow newRow;

                            for (int i = 0; i < end - begin; i++) {
                                // 插入数据的序列号
                                String serialNumber = per + (begin + i + 1);
                                // excel插入一行，下方内容下移一行
                                if (rowNum + i + 1 <= sheet.getLastRowNum())
                                {
                                    sheet.shiftRows(rowNum + i + 1, sheet.getLastRowNum(), 1);
                                }
                                newRow = sheet.createRow(rowNum + i + 1);
                                copyRow(row, newRow);

                                // 修改序列号，序号和数量不改，用于核对数据
                                newRow.getCell(BCellNum).setCellValue(serialNumber);
                                newRow.getCell(ECellNum).setCellValue(serialNumber);

                                System.out.println("新增数据：" + serialNumber);
                            }
                        }
                    }
                }
            }
            saveExcel(workbook);
            workbook.close();
            is.close();
        }
    }

    /**
     * 行复制功能
     *
     * @param fromRow
     * @param toRow
     */
    public static void copyRow(XSSFRow fromRow, XSSFRow toRow) {
        for (Iterator cellIt = fromRow.cellIterator(); cellIt.hasNext(); ) {
            XSSFCell tmpCell = (XSSFCell) cellIt.next();
            XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());
            copyCell(tmpCell, newCell, true);
        }
    }

    /**
     * 复制单元格
     *
     * @param srcCell
     * @param distCell
     * @param copyValueFlag true则连同cell的内容一起复制
     */
    public static void copyCell(XSSFCell srcCell, XSSFCell distCell, boolean copyValueFlag) {
        //样式
        distCell.setCellStyle(srcCell.getCellStyle());
        //评论
        if (srcCell.getCellComment() != null) {
            distCell.setCellComment(srcCell.getCellComment());
        }
        //复制值
        if (copyValueFlag) {
            distCell.setCellValue(MenuUtil.getCellValue(srcCell));
        }
    }

    /**
     * 保存工作薄
     * @param wb
     */
    public static void saveExcel(XSSFWorkbook wb) {
        FileOutputStream fileOut;
        try {
            fileOut = new FileOutputStream(excelPath);
            wb.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
