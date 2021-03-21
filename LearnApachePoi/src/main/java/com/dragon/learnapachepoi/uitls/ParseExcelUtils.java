package com.dragon.learnapachepoi.uitls;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 10:43
 * @description：
 * @modified By：
 * @version: $
 */
public class ParseExcelUtils {
    private static final Logger logger = LoggerFactory.getLogger(ParseExcelUtils.class);

    /**
     * 解析.xlsx excel
     * @param fis
     * @return
     */
    public static List<List<String>> parseExcelxlsx(InputStream fis) {
        List<List<String>> data = new ArrayList<List<String>>();
        try {
            @SuppressWarnings("resource")
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int firstRow = sheet.getFirstRowNum();// 第一行序号
            int lastRow = sheet.getLastRowNum();// 最后一行序号

            for (int i = firstRow + 1; i < lastRow + 1; i++) {
                XSSFRow row = sheet.getRow(i); // 获取第行

                List<String> map = new ArrayList<String>();
                int firstCell = 0;
                int lastCell = row.getLastCellNum();// 最后一列序号
                for (int j = firstCell; j < lastCell; j++) {
                    XSSFCell cell = row.getCell(j);// 当前单元格
                    // 当单元格是数字格式时，需要把它的cell type转成String，否则会出错
                    // 当前单元格的值
                    // cell.setCellType(Cell.CELL_TYPE_STRING);//
                    // 当前cell为空格的时候，格式为数字类型，可能出现cell为null值情况，这儿也会直接抛出异常错误。导致解析失败！正常应该直接读取当前cell为空字符填充就可以
                    String val = "";
                    if (cell != null) {
                        cell.setCellType(CellType.STRING);// 2020-03-19
                        val = cell.getStringCellValue();
                    }
                    map.add(val);
                }
                data.add(map);
            }
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            throw new RuntimeException("导入异常");
        }
        return data;
    }

    /**
     *  解析.xls excel
     * @param fis
     * @return
     */
    public static List<List<String>> parseExcelxls(InputStream fis) {
        List<List<String>> data = new ArrayList<List<String>>();
        try {
            @SuppressWarnings("resource")
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);
            int firstRow = sheet.getFirstRowNum();// 第一行序号
            int lastRow = sheet.getLastRowNum();// 最后一行序号

            for (int i = firstRow + 1; i < lastRow + 1; i++) {
                HSSFRow row = sheet.getRow(i); // 获取第行

                List<String> map = new ArrayList<String>();
                int firstCell = 0;
                int lastCell = row.getLastCellNum();// 最后一列序号
                for (int j = firstCell; j < lastCell; j++) {
                    HSSFCell cell = row.getCell(j);// 当前单元格
                    // 当单元格是数字格式时，需要把它的cell type转成String，否则会出错
                    // 当前单元格的值
                    // cell.setCellType(Cell.CELL_TYPE_STRING);
                    String val = "";
                    if (cell != null) {
                        cell.setCellType(CellType.STRING);
                        val = cell.getStringCellValue();
                    }
                    map.add(val);
                }
                data.add(map);
            }
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            throw new RuntimeException("导入异常");
        }
        return data;
    }


    /**
     * 解析.xlsx excel ：拓展（处理数字格式时精度丢失问题）
     * @param fis
     * @return
     */
    public static List<List<String>> parseExcelXlsxForMult(InputStream fis) {
        List<List<String>> data = new ArrayList<List<String>>();
        try {
            @SuppressWarnings("resource")
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int firstRow = sheet.getFirstRowNum();// 第一行序号
            int lastRow = sheet.getLastRowNum();// 最后一行序号

            for (int i = firstRow + 1; i < lastRow + 1; i++) {
                XSSFRow row = sheet.getRow(i); // 获取第行

                List<String> map = new ArrayList<String>();
                int firstCell = 0;
                int lastCell = row.getLastCellNum();// 最后一列序号
                for (int j = firstCell; j < lastCell; j++) {
                    XSSFCell cell = row.getCell(j);// 当前单元格
                    // 当单元格是数字格式时，需要把它的cell type转成String，否则会出错
                    // 当前单元格的值
                    // cell.setCellType(Cell.CELL_TYPE_STRING);//
                    // 当前cell为空格的时候，格式为数字类型，可能出现cell为null值情况，这儿也会直接抛出异常错误。导致解析失败！正常应该直接读取当前cell为空字符填充就可以
                    String val = "";
                    if (cell != null) {
                        CellType cellTypeEnum = cell.getCellTypeEnum();
                        if(cell.getCellTypeEnum() == CellType.NUMERIC){// 数字类型，解决解析出的 数字精度丢失问题
                            val = realStringValueOfDouble(cell.getNumericCellValue());
                        }else{
                            cell.setCellType(CellType.STRING);// 2020-03-19
                            val = cell.getStringCellValue();
                        }
                    }
                    map.add(val);
                }
                data.add(map);
            }
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            throw new RuntimeException("导入异常");
        }
        return data;
    }

    /**
     *  解析.xls ：拓展（处理数字格式时精度丢失问题）
     * @param fis
     * @return
     */
    public static List<List<String>> parseExcelxlsForMult(InputStream fis) {
        List<List<String>> data = new ArrayList<List<String>>();
        try {
            @SuppressWarnings("resource")
            HSSFWorkbook workbook = new HSSFWorkbook(fis);
            HSSFSheet sheet = workbook.getSheetAt(0);
            int firstRow = sheet.getFirstRowNum();// 第一行序号
            int lastRow = sheet.getLastRowNum();// 最后一行序号

            for (int i = firstRow + 1; i < lastRow + 1; i++) {
                HSSFRow row = sheet.getRow(i); // 获取第行

                List<String> map = new ArrayList<String>();
                int firstCell = 0;
                int lastCell = row.getLastCellNum();// 最后一列序号
                for (int j = firstCell; j < lastCell; j++) {
                    HSSFCell cell = row.getCell(j);// 当前单元格
                    // 当单元格是数字格式时，需要把它的cell type转成String，否则会出错
                    // 当前单元格的值
                    // cell.setCellType(Cell.CELL_TYPE_STRING);
                    String val = "";
                    if (cell != null) {
                        if(cell.getCellTypeEnum() == CellType.NUMERIC){// 数字类型，解决解析出的 数字精度丢失问题
                            val = realStringValueOfDouble(cell.getNumericCellValue());
                        }else{
                            cell.setCellType(CellType.STRING);// 2020-03-19
                            val = cell.getStringCellValue();
                        }
                    }
                    map.add(val);
                }
                data.add(map);
            }
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            throw new RuntimeException("导入异常");
        }
        return data;
    }

    /**
     * Excel 读取文件精度问题
     * @param d
     * @return
     */
    public static String realStringValueOfDouble(Double d) {
        String doubleStr = d.toString();
        boolean b = doubleStr.contains("E");
        int indexOfPoint = doubleStr.indexOf('.');
        if (b) {
            int indexOfE = doubleStr.indexOf('E');
            BigInteger xs = new BigInteger(doubleStr.substring(indexOfPoint + BigInteger.ONE.intValue(), indexOfE));
            int pow = Integer.parseInt(doubleStr.substring(indexOfE + BigInteger.ONE.intValue()));
            int xsLen = xs.toByteArray().length;
            int scale = xsLen - pow > 0 ? xsLen - pow : 0;
            final String format = "%." + scale + "f";
            doubleStr = String.format(format, d);
        } else {
            java.util.regex.Pattern p = Pattern.compile(".0$");
            java.util.regex.Matcher m = p.matcher(doubleStr);
            if (m.find()) {
                doubleStr = doubleStr.replace(".0", "");
            }
        }
        return doubleStr;
    }

}
