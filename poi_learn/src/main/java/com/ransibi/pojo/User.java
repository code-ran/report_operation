package com.ransibi.pojo;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.util.Date;
import java.util.List;


@Data
public class User {
    private Long id;
    /**
     * 员工名
     */
    private String userName;
    /**
     * 手机号
     */
    private String phone;
    /**
     * 省份名
     */
    private String province;
    /**
     * 城市名
     */
    private String city;
    /**
     * 工资
     */
    private Integer salary;
    /**
     * 入职日期
     */
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date hireDate;

    private String hireDateFormat;
    /**
     * 部门id
     */
    private Integer deptId;
    /**
     * 出生日期
     */
    private Date birthday;
    private String birthdayFormat;
    /**
     * 一寸照片
     */
    private String photo;
    /**
     * 现在居住地址
     */
    private String address;
    /**
     * 办公用品
     */
    private List<Resource> resourceList;

}
