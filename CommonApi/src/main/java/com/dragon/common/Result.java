package com.dragon.common;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 10:30
 * @description：返回值封装
 * @modified By：
 * @version: $
 */
public class Result {

    public static final String SUCCESS = "success";
    public static final String ERROR = "error";
    public static final String WARN = "warn";
    public static final String FAILURE = "failure";
    public String state = null;
    public String message = null;
    public Object data = null;

    public static Result SUCCESS() {
        Result result = new Result();
        result.state = "success";
        return result;
    }

    public Result() {
    }

    public String getSUCCESS() {
        return "success";
    }

    public String getERROR() {
        return "error";
    }

    public String getWARN() {
        return "warn";
    }

    public String getState() {
        return this.state;
    }

    public String getMessage() {
        return this.message;
    }

    public Object getData() {
        return this.data;
    }

    public void setState(String state) {
        this.state = state;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public void setData(Object data) {
        this.data = data;
    }
}
