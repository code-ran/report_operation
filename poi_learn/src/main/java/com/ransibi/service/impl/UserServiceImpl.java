package com.ransibi.service.impl;

import com.ransibi.dao.UserMapper;
import com.ransibi.service.IUserService;
import org.apache.commons.collections4.CollectionUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.*;


@Service
public class UserServiceImpl implements IUserService {

    @Autowired
    UserMapper userMapper;

    @Override
    public String getUserInfo() {
        List<Map<String, Object>> mapList = userMapper.selectUserInfo();
        List<String> ptIdAll = new ArrayList<>(Arrays.asList("1","2","3","4"));
        List<String> existPtId= new ArrayList<>(Arrays.asList("1","2"));
        Collection<String> subtract = CollectionUtils.subtract(ptIdAll, existPtId);
        System.out.println(subtract.toString());
        return mapList.toString();
    }
}
