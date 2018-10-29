package indi.liht.stat.utils;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;

/**
 * Usage:
 * Poi操作MS Office文档工具类
 * @author lihongtao ibraxwell@sina.com
 * on 2018/10/28
 **/
public class PoiUtils {

    /**
     * 根据后缀，使用不同方法从工作表重获取工作簿
     * @param excelFilePath Excel完整路径
     * @param sheetNames 工作簿名
     * @return Sheet
     */
    public static Sheet[] getSheetsFromPath(String excelFilePath, String... sheetNames) {
        if (excelFilePath.toLowerCase().endsWith(".xls")) {
            return getHSSFSheetsFromPath(excelFilePath, sheetNames);
        } else if (excelFilePath.toLowerCase().endsWith("xlsx")) {
            return getXSSFSheetsFromPath(excelFilePath, sheetNames);
        } else {
            return null;
        }
    }

    /**
     * 返回.xls后缀，名为sheetName的工作簿
     * @param excelFilePath Excel（.xls后缀）完整路径
     * @param sheetNames 工作簿名
     * @return Sheet
     */
    private static HSSFSheet[] getHSSFSheetsFromPath(String excelFilePath, String... sheetNames) {
        HSSFSheet[] sheet = new HSSFSheet[sheetNames.length];
        InputStream is = null;
        HSSFWorkbook workbook = null;
        try {
            is = new FileInputStream(excelFilePath);
            workbook = new HSSFWorkbook(is);
            for (int i = 0, size = sheetNames.length; i < size; i++) {
                // 通过sheetName获取Sheet
                sheet[i] = workbook.getSheet(sheetNames[i]);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(workbook);
            IOUtils.closeQuietly(is);
        }
        System.out.println("加载[" + excelFilePath + "]工作簿[类型：OLE2]的"
                + Arrays.toString(sheetNames) + "工作表完成！");
        return sheet;
    }

    /**
     * 返回.xlsx后缀，名为sheetName的工作簿
     * @param excelFilePath Excel（.xlsx）完整路径
     * @param sheetNames 工作簿名
     * @return Sheet
     */
    private static XSSFSheet[] getXSSFSheetsFromPath(String excelFilePath, String... sheetNames) {
        XSSFSheet[] sheet = new XSSFSheet[sheetNames.length];
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(excelFilePath);
            for (int i = 0, size = sheetNames.length; i < size; i++) {
                // 通过sheetName获取Sheet
                sheet[i] = workbook.getSheet(sheetNames[i]);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(workbook);
        }
        System.out.println("加载[" + excelFilePath + "]工作簿[类型：OOXML]的"
                + Arrays.toString(sheetNames) + "工作表完成！");
        return sheet;
    }

    /**
     * 根据cell的cellType类型判断，调用不同方法取值（只判断了字符串和数字类型的）
     * @param cell 要取值的cell
     * @return cell的值
     */
    public static Object getValueFromCell(Cell cell) {
        if (cell == null) {
            return null;
        }
        Object value;
        switch (cell.getCellType()) {
            case NUMERIC:
                value = cell.getNumericCellValue();
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            default:
                value = null;
        }
        return value;
    }

}
