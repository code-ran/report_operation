package com.ransibi;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;


@SpringBootApplication
@MapperScan("com.ransibi.dao")
public class Main {
    public static void main(String[] args) {
        SpringApplication.run(Main.class);
    }
}