package com.ransibi.pojo;

import com.alibaba.fastjson.JSONArray;
import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

@Data
public class ReCloseBaseBean {
    private Integer id;
    /**
     * 地区名称
     */
    private String areaName;
    /**
     * 厂站名称
     */
    private String stnName;
    /**
     * 装置名称
     */
    private String ptName;
    private String ptId;
    /**
     * 是否告警:0-不判断,1-正常,2-异常
     */
    private Integer isAlarm;
    /**
     * 不输出告警原因
     */
    private Integer reason;
    /**
     * 投退情况(基准)
     */
    private Integer standardValue;
    /**
     * 投退情况
     */
    private Integer status;
    /**
     * 投退一致情况
     */
    private Integer orderCk;
    /**
     * 比对时间
     */
    @JsonFormat(pattern = "yyyy-MM-dd HH:mm:ss")
    private Date chkTime;
    /**
     * 数据源
     */
    private List<String> tableDataSourceLst;

    private JSONArray sfDiInfo;
    //charge query use

    private String sgInfo;

    private String sfInfo;

    private String diInfo;


    //charge export use

    private String sgName;

    private String sgValue;

    private String sfName;

    private String sfValue;

    private String diName;

    private String diValue;

    private String isAlarmName;

    private String chkTimeFormat;

    public String getChkTimeFormat() {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        chkTimeFormat = dateFormat.format(chkTime);
        return chkTimeFormat;
    }

    public void setChkTimeFormat(String chkTimeFormat) {
        this.chkTimeFormat = chkTimeFormat;
    }
}
