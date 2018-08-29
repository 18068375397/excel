package excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DealExcelDateUtil {

    // 1、当前文件已经存在
    private static String excelPath = "C:/Users/asus/Desktop/deal.xlsx";
    // 2、设置为文本格式
    // 3、起始序列号所在单元格(从0开始)
    private static int BCellNum = 5;
    private static int ECellNum = 6;

    public static void main(String[] args) throws Exception {
        dealExcel();
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
                // 只处理第一个sheet
                for (int sheetNum = 0; sheetNum < 1; sheetNum++) {
                    // 获得当前sheet工作表
                    XSSFSheet sheet = workbook.getSheetAt(sheetNum);
                    if (sheet == null) {
                        continue;
                    }
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

                        int begin = 0;
                        int end = 0;

                        // 序列号不同，需要处理数据并插入
                        if (!beginNumber.equals(endNumber))
                        {
                            String serStart = getDigit(beginNumber);
                            String serEnd = getDigit(endNumber);
                            int length = serStart.length();

                            begin = Integer.valueOf(serStart);
                            end = Integer.valueOf(serEnd);
                            cell2.setCellValue(beginNumber);
                            System.out.println("第" + (rowNum + 1) + "行数据处理，共" + (end - begin) + "条。起始序列号为：" + beginNumber + "，结束序列号为：" + endNumber);

                            // 复制所需插入数据
                            XSSFRow newRow;

                            for (int i = 0; i < end - begin; i++) {
                                // 提取唯一码前缀
                                String per = getLetter(beginNumber);
                                String serialNumber = per + String.format("%0" + length + "d", begin + i + 1);

                                // excel插入一行，下方内容下移一行
                                if (rowNum + i + 1 <= sheet.getLastRowNum()) {
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
     *
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

    //提取字母
    public static String getLetter(String a) {
        StringBuffer sb = new StringBuffer();
        for (int i = 0; i < a.length(); i++) {
            char c = a.charAt(i);

            if ((c <= 'z' && c >= 'a') || (c <= 'Z' && c >= 'A')) {
                sb.append(c);
            }
        }
        System.out.println(sb.toString());
        return sb.toString();
    }

    //提取数字
    public static String getDigit(String text) {
        String digitList = "";
        for (int i = 0; i < text.length(); i++) {
            int j = (int) text.charAt(i);
            if ((j > 47 && j < 58)) {
                digitList = digitList + (j - 48);
            }
        }
        return digitList;
    }

}
