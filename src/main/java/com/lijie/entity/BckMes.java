package com.lijie.entity;

import com.alibaba.fastjson.JSONObject;

import java.io.Serializable;

public class BckMes implements Serializable {

    /**
     * 返回请求状态 0 失败 1 成功
     */
    private int status;
    /**
     * 请求返回信息
     */
    private String msg;
    /**
     * 返回数据
     */
    private Object data;

    public BckMes(int status, String msg, Object data) {
        this.status = status;
        this.msg = msg;
        this.data = data;
    }

    public static <T> String successMes(T data) {
        return successMes("success", data);
    }

    public static <T> String successMes(String msg, T data) {
        BckMes bckMes = new BckMes(1, msg, data);
        return bckMes.toString();
    }

    public static <T> String errorMes() {
        return errorMes("error");
    }

    public static <T> String errorMes(String mes) {
        BckMes bckMes = new BckMes(0, mes, "");
        return bckMes.toString();
    }


    @Override
    public String toString() {
        return JSONObject.toJSONString(this);
    }

    public int getStatus() {
        return status;
    }

    public void setStatus(int status) {
        this.status = status;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }
}
