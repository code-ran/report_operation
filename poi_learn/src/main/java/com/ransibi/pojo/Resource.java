package com.ransibi.pojo;

import lombok.Data;

/**
 * 员工领取的办公用品记录表
 */
@Data
public class Resource {

    private Long id;
    /**
     * 用品名称
     */
    private String name;
    /**
     * 价格
     */
    private Double price;
    /**
     * 员工id
     */
    private Long userId;
    /**
     * 是否需要归还
     */
    private Boolean needReturn;
    /**
     * 照片
     */
    private String photo;
}
