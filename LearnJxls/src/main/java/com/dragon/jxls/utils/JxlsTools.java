package com.dragon.jxls.utils;

import org.apache.commons.lang3.StringUtils;
import org.jxls.transform.poi.WritableCellValue;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 14:06
 * @description：jxls 模板编辑用工具类
 * @modified By：
 * @version: $
 */
public class JxlsTools {
    private final static Logger logger = LoggerFactory.getLogger(JxlsTools.class);

    /**
     * 日期格式化
     *
     * @param fmt
     * @return
     */
    public String dateFmt(String fmt) {
        Date dateTemp = new Date();
        if (fmt == null) {
            return "";
        }
        try {
            SimpleDateFormat dateFmt = new SimpleDateFormat(fmt);
            String dateFmtStr = dateFmt.format(dateTemp);
            return dateFmtStr;
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
        }
        return "";
    }

    /**
     * 传入data类型的转化
     * @param fmt
     * @return
     */
    public String formatDate(Date dateTemp, String fmt) {
        if (fmt == null) {
            return "";
        }
        try {
            SimpleDateFormat dateFmt = new SimpleDateFormat(fmt);
            String dateFmtStr = dateFmt.format(dateTemp);
            return dateFmtStr;
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
        }
        return "";
    }

    /**
     * if判断
     *
     * @param b
     * @param o1
     * @param o2
     * @return
     */
    public Object ifElse(boolean b, Object o1, Object o2) {
        return b ? o1 : o2;
    }

    /**
     * 格式化数字
     * @param oTr
     * @param lengthStr
     * @return
     */
    public String formatNumber(String oTr, String lengthStr) {
        String tempStr = "";
        if (StringUtils.isNotBlank(oTr) && StringUtils.isNotBlank(lengthStr)) {
            BigDecimal fmtNum = new BigDecimal(oTr);
            int numInt = Integer.parseInt(lengthStr);// 保留小数位长度
            String conStr = "";
            for (int i = 0; i < numInt; i++) {
                conStr = conStr + "0";
            }

            DecimalFormat df4 = new DecimalFormat("#,##0." + conStr);
            tempStr = df4.format(fmtNum);
        }
        return tempStr;
    }

    /**
     * 连接字符
     *
     * @param str1
     * @param str2
     * @return
     */
    public String connectStr(String str1, String str2) {
        if (StringUtils.isBlank(str1) && !StringUtils.isBlank(str2)) {
            return str2;
        }
        if (StringUtils.isBlank(str2) && !StringUtils.isBlank(str1)) {
            return str1;
        }
        return str1 + str2;
    }

    /**
     * 数字类型转化为字符类型
     * @param number
     * @return
     */
    public String numberToStr(BigDecimal number) {
        if (number != null) {
            return number.toString();
        }
        return "";
    }

    /**
     * 字符类型转为数字类型
     * @param str
     * @return
     */
    public BigDecimal strToNumber(String str) {
        if (StringUtils.isNotBlank(str)) {
            return new BigDecimal(str);

        }
        return null;
    }

    /**
     *
     * @param dataStr
     * @param format
     * @return
     */
    public static String getDateWithFormat(String dataStr,String format){
        String returnString = "";
        SimpleDateFormat sdf1 = new SimpleDateFormat(format);
        SimpleDateFormat sdf2 = new SimpleDateFormat("yyyyMMdd");
        try {
            Date date = sdf2.parse(dataStr);
            returnString = sdf1.format(date);
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
        }
        return returnString;
    }

    /**
     * 数字保留几位小数
     * @param oStr
     * @param lengthStr
     * @return
     */
    public String formatRatioNum(String oStr, String lengthStr) {
        String tempStr = "";
        if (StringUtils.isNotBlank(oStr) && StringUtils.isNotBlank(lengthStr)) {
            BigDecimal fmtNum = new BigDecimal(oStr);
            int numInt = Integer.parseInt(lengthStr);
            String conStr = "";
            for (int i = 0; i < numInt; i++) {
                conStr = conStr + "0";
            }
            DecimalFormat df4 = new DecimalFormat("#." + conStr);
            tempStr = df4.format(fmtNum);
        }
        return tempStr;
    }

    /**
     * 获取集合中对象
     *
     * @param list
     * @param lengthStr
     * @return
     */
    public <E> E getList(List<E> list, String lengthStr) {
        if (list != null && list.size() > 0) {
            int numInt = Integer.parseInt(lengthStr);
            return list.get(numInt);
        }
        return null;
    }

    public static String formatPercent(String objStr) {
        String tempStr = "";
        if (StringUtils.isNotBlank(objStr)) {
            BigDecimal fmtNum = new BigDecimal(objStr);
            BigDecimal multiply = fmtNum.multiply(new BigDecimal(100)).setScale(2,BigDecimal.ROUND_HALF_UP);
            tempStr = multiply + "%";
        }
        return tempStr;
    }

    //生成下拉菜单07版
    public WritableCellValue dropList(String splitItem , String value){
        return new DropdownCellValue(splitItem, value);
    }
    //生成下拉菜单03版
    public WritableCellValue dropListForXls(String splitItem , String value){
        return new DropdownCellValueForXls(splitItem, value);
    }

    public static void main(String[] args) {
        String str = formatPercent("0.9");
        logger.info(str);
    }

}
