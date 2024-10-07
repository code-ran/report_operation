package com.ransibi.service;


import com.ransibi.pojo.User;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

public interface IUserService {
    List<User> getUserInfo(Integer page, Integer pageSize);

    String uploadExcelInfo(MultipartFile file) throws Exception;

    void downLoadXlsxByPoi(HttpServletResponse response) throws Exception;

    void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws Exception;

    void downLoadXlsxWithTemplate(HttpServletResponse response) throws Exception;

    void exportTest(HttpServletResponse response) throws Exception;
}
