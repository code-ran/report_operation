package com.ransibi.controller;

import com.ransibi.pojo.User;
import com.ransibi.service.ITestService;
import com.ransibi.service.IUserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

@RestController
@RequestMapping("user")
public class UserController {

    @Autowired
    IUserService iUserService;

    @Autowired
    ITestService iTestService;

    @GetMapping("/findPage")
    public List<User> getUser(
            @RequestParam(value = "page",defaultValue = "1") Integer page,
            @RequestParam(value = "rows",defaultValue = "10") Integer pageSize){
        return iUserService.getUserInfo(page,pageSize);
    }

    @PostMapping("/uploadExcel")
    public String uploadExcel(MultipartFile file) throws Exception {
        return iUserService.uploadExcelInfo(file);
    }

    @GetMapping("/downLoadXlsxByPoi")
    public void downLoadXlsxByPoi( HttpServletResponse response) throws Exception{
        //无样式导出
//        iUserService.downLoadXlsxByPoi(response);
        //含样式导出
//        iUserService.downLoadXlsxByPoiWithCellStyle(response);
        //通过模版导出
//        iUserService.downLoadXlsxWithTemplate(response);
        //个人测试案例导出
        iTestService.exportTest(response);
    }

}
