package com.dragon.common;

import java.util.HashMap;
import java.util.Map;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 11:13
 * @description：查询参数封装
 * @modified By：
 * @version: $
 */
public class SearchParam {

    private int start = 0;

    private int limit = 20;

    private int total = 0;

    private int pageNo = 0;

    public int getEnd() {
        return start + limit;
    }

    // 在拦截器中是否自动计算count值
    public boolean isAutoTotal = true;

    private Map<String, Object> sp = new HashMap<String, Object>();

    public int getStart() {
        start = (pageNo - 1) * limit;
        return start;
    }

    public void setStart(int start) {
        this.start = start;
    }

    public int getLimit() {
        return limit;
    }

    public void setLimit(int limit) {
        this.limit = limit;
    }

    public Map<String, Object> getSp() {
        return sp;
    }

    public void setSp(Map<String, Object> sp) {
        this.sp = sp;
    }

    public int getTotal() {
        return total;
    }

    public void setTotal(int total) {
        this.total = total;
    }

    public int getPageNo() {
        return pageNo;
    }

    public void setPageNo(int pageNo) {
        this.pageNo = pageNo;
    }

    @Override
    public String toString() {
        String tempString = "{";
        for (String key : this.sp.keySet()) {
            if(sp.get(key) != null && !"".equals(sp.get(key)) && "{".equals(tempString)){
                tempString = tempString + key+"="+ sp.get(key);
            }else if(sp.get(key) != null && !"".equals(sp.get(key))){
                tempString = tempString +", "+key+"="+ sp.get(key);
            }
        }
        return tempString+", pageNo="+pageNo + ", limit="+limit + "}";
    }

    /**
     * 在拦截器中是否自动计算count值
     * @return
     */
    public boolean isAutoTotal() {
        return isAutoTotal;
    }

    public void setAutoTotal(boolean isAutoTotal) {
        this.isAutoTotal = isAutoTotal;
    }

}
