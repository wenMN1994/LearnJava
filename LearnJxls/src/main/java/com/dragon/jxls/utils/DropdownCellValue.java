package com.dragon.jxls.utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.jxls.common.Context;
import org.jxls.transform.poi.WritableCellValue;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 14:01
 * @description：自定义下拉框07版
 * @modified By：
 * @version: $
 */
public class DropdownCellValue implements WritableCellValue {

    private String splitItem , value;

    public DropdownCellValue(String splitItem , String value){
        this.splitItem = splitItem;
        this.value = value;
    }

    @Override
    public Object writeToCell(Cell cell, Context context) {
        String[] downs = splitItem.split(",");
        XSSFSheet sheet = (XSSFSheet) cell.getSheet();
        //开始行、结束行、开始列、结束列
        Row row = cell.getRow();
        int startRowIndex = row.getRowNum();
        int endRowIndex = startRowIndex;
        int cellStartIndex = cell.getColumnIndex();
        int cellEndIndex = cellStartIndex;
        CellRangeAddressList regions = new CellRangeAddressList(startRowIndex, endRowIndex, cellStartIndex, cellEndIndex);
        DataValidationHelper dataValidationHelper = sheet.getDataValidationHelper();
        DataValidationConstraint createExplicitListConstraint = dataValidationHelper.createExplicitListConstraint(downs);
        DataValidation createValidation = dataValidationHelper.createValidation(createExplicitListConstraint, regions);
        //处理Excel兼容性问题
        if (createValidation instanceof XSSFDataValidation) {
            createValidation.setSuppressDropDownArrow(true);
            createValidation.setShowErrorBox(true);
        } else {
            createValidation.setSuppressDropDownArrow(false);
        }
        sheet.addValidationData(createValidation);
        cell.setCellValue(value);
        return cell;
    }
}
