package com.dragon.learnapachepoi.poi.uitls;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.awt.Color;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.List;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 10:33
 * @description：
 * @modified By：
 * @version: $
 */
public class ExportExcelUtils {

    private static Logger logger = LoggerFactory.getLogger(ExportExcelUtils.class);

    public static void exportExcel(HttpServletResponse response, String fileName, ExcelData data) throws Exception {
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/vnd.ms-excel");
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename="+URLEncoder.encode(fileName, "utf-8"));
        exportExcel(data, response.getOutputStream());
    }

    public static void exportExcel(ExcelData data, OutputStream out) throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();
        try {
            String sheetName = data.getName();
            if (null == sheetName) {
                sheetName = "Sheet1";
            }
            XSSFSheet sheet = wb.createSheet(sheetName);
            writeExcel(wb, sheet, data);

            wb.write(out);
        } finally {
            if(wb != null){
                try {
                    wb.close();
                } catch (Exception e) {
                    logger.error(e.getMessage(),e);
                }
            }
        }
    }

    //多工作表导出
    public static void exportExcelMultiSheet(HttpServletResponse response, String fileName, List<ExcelData> dataList) throws Exception {
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/vnd.ms-excel");
        OutputStream out = response.getOutputStream();
        XSSFWorkbook wb = new XSSFWorkbook();
        int i = 1;
        try {
            for (ExcelData ed : dataList) {
                // 下载文件的默认名称
                response.setHeader("Content-Disposition", "attachment;filename="+URLEncoder.encode(fileName, "utf-8"));
                String sheetName = ed.getName();
                if (null == sheetName) {
                    sheetName = "Sheet"+i;
                }
                XSSFSheet sheet = wb.getSheet(sheetName);
                if(sheet==null) {
                    sheet = wb.createSheet(sheetName);
                }
                writeExcel(wb, sheet, ed);
                i++;
            }
            wb.write(out);
        } finally {
            if(wb != null){
                try {
                    wb.close();
                } catch (Exception e) {
                    logger.error(e.getMessage(),e);
                }
            }
        }
    }

    public static void writeExcel(XSSFWorkbook wb, Sheet sheet, ExcelData data) {
        // int rowIndex = 0;
        int rowIndex = sheet.getLastRowNum();
        if(rowIndex>0) {
            rowIndex+=2;
        }
        rowIndex = writeTitlesToExcel(wb, sheet, data.getTitles(),rowIndex);
        writeRowsToExcel(wb, sheet, data.getRows(), rowIndex);
        autoSizeColumns(sheet, data.getTitles().size() + 1);
    }

    public static int writeTitlesToExcel(XSSFWorkbook wb, Sheet sheet, List<String> titles,int rowIndex) {
        // int rowIndex = 0;
        int colIndex = 0;

        Font titleFont = wb.createFont();
        titleFont.setFontName("simsun");
        ((XSSFFont) titleFont).setBold(true);
        // titleFont.setFontHeightInPoints((short) 14);
        titleFont.setColor(IndexedColors.BLACK.index);

        XSSFCellStyle titleStyle = wb.createCellStyle();
        titleStyle.setAlignment(titleStyle.getAlignmentEnum().CENTER  );
        titleStyle.setVerticalAlignment(titleStyle.getVerticalAlignmentEnum().CENTER);
        titleStyle.setFillForegroundColor(new XSSFColor(new Color(182, 184, 192)));
        titleStyle.setFillPattern(titleStyle.getFillPatternEnum().SOLID_FOREGROUND );
        titleStyle.setFont(titleFont);
        setBorder(titleStyle, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));

        XSSFCellStyle dataStyle = wb.createCellStyle();
        XSSFDataFormat dataFormat = wb.createDataFormat();
        dataStyle.setDataFormat(dataFormat.getFormat("@"));

        Row titleRow = sheet.createRow(rowIndex);
        // titleRow.setHeightInPoints(25);
        colIndex = 0;

        for (String field : titles) {
            sheet.setDefaultColumnStyle(colIndex, dataStyle);
            Cell cell = titleRow.createCell(colIndex);
            cell.setCellValue(field);
            cell.setCellStyle(titleStyle);
            colIndex++;
        }

        rowIndex++;
        return rowIndex;
    }

    public static int writeRowsToExcel(XSSFWorkbook wb, Sheet sheet, List<List<Object>> rows, int rowIndex) {
        int colIndex = 0;
        Font dataFont = wb.createFont();
        dataFont.setFontName("simsun");
        // dataFont.setFontHeightInPoints((short) 14);
        dataFont.setColor(IndexedColors.BLACK.index);

        XSSFCellStyle dataStyle = wb.createCellStyle();
        dataStyle.setAlignment(dataStyle.getAlignmentEnum().CENTER );
        dataStyle.setVerticalAlignment(dataStyle.getVerticalAlignmentEnum().CENTER);
        dataStyle.setFont(dataFont);
        setBorder(dataStyle, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));

        for (List<Object> rowData : rows) {
            Row dataRow = sheet.createRow(rowIndex);
            // dataRow.setHeightInPoints(25);
            colIndex = 0;

            for (Object cellData : rowData) {
                Cell cell = dataRow.createCell(colIndex);
                if (cellData != null) {
                    cell.setCellValue(cellData.toString());
                } else {
                    cell.setCellValue("");
                }

                cell.setCellStyle(dataStyle);
                colIndex++;
            }
            rowIndex++;
        }
        return rowIndex;
    }



    public static void autoSizeColumns(Sheet sheet, int columnNumber) {

        for (int i = 0; i < columnNumber; i++) {
            int orgWidth = sheet.getColumnWidth(i);
            sheet.autoSizeColumn(i, true);
            int newWidth = (int) (sheet.getColumnWidth(i) + 100);
            if (newWidth > orgWidth) {
                sheet.setColumnWidth(i, newWidth);
            } else {
                sheet.setColumnWidth(i, orgWidth);
            }
        }
    }

    public static void setBorder(XSSFCellStyle style, BorderStyle border, XSSFColor color) {
        style.setBorderTop(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);
        style.setBorderBottom(border);
        style.setBorderColor(BorderSide.TOP, color);
        style.setBorderColor(BorderSide.LEFT, color);
        style.setBorderColor(BorderSide.RIGHT, color);
        style.setBorderColor(BorderSide.BOTTOM, color);
    }
}
