package com.ransibi.dao;

import com.github.pagehelper.Page;
import com.ransibi.pojo.User;
import org.springframework.stereotype.Repository;

import java.util.List;


@Repository
public interface UserMapper {
    List<User> selectUserInfo();

    void insertUser(List<User> userList);
}
