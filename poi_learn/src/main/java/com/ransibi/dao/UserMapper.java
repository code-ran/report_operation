package com.ransibi.dao;

import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Map;


@Repository
public interface UserMapper {
    List<Map<String,Object>> selectUserInfo();
}
