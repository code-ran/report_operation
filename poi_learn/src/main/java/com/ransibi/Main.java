package com.ransibi;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;


@SpringBootApplication
@MapperScan("com.ransibi.dao")
public class Main {
    static {
        //程序获取数据库连接，对已存在的数据库连接进行检查，检查到空闲时间过久的连接会进行注销，并报出错误提示。
        System.setProperty("druid.mysql.usePingMethod", "false");
    }

    public static void main(String[] args) {
        SpringApplication.run(Main.class);
    }
}