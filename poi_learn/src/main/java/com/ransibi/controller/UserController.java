package com.ransibi.controller;

import com.ransibi.service.IUserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("rsb/user")
public class UserController {

    @Autowired
    IUserService iUserService;

    @GetMapping("/list")
    public String getUser(){
        return iUserService.getUserInfo();
    }
}
