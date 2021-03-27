package com.dragon.jxls.utils;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jxls.common.Context;
import org.jxls.transform.poi.WritableCellValue;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 14:02
 * @description：自定义下拉框03版
 * @modified By：
 * @version: $
 */
public class DropdownCellValueForXls implements WritableCellValue {

    private String splitItem, value;

    public DropdownCellValueForXls(String splitItem, String value) {
        this.splitItem = splitItem;
        this.value = value;
    }

    @Override
    public Object writeToCell(Cell cell, Context context) {
        String[] downs = splitItem.split(",");
        DVConstraint dv = DVConstraint.createExplicitListConstraint(downs);
        Sheet sheet = cell.getSheet();
        // 开始行、结束行、开始列、结束列
        Row row = cell.getRow();
        int startRowIndex = row.getRowNum();
        int endRowIndex = startRowIndex;
        int cellStartIndex = cell.getColumnIndex();
        int cellEndIndex = cellStartIndex;
        CellRangeAddressList regions = new CellRangeAddressList(startRowIndex, endRowIndex, cellStartIndex,
                cellEndIndex);
        DataValidation v = new HSSFDataValidation(regions, dv);
        sheet.addValidationData(v);
        cell.setCellValue(value);
        return cell;

    }

}
