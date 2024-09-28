package com.ransibi.service;


import com.ransibi.pojo.User;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

public interface IUserService {
    List<User> getUserInfo(Integer page, Integer pageSize);

    String uploadExcelInfo(MultipartFile file) throws Exception;

    void downLoadXlsxByPoiWithTemplate(HttpServletResponse response) throws Exception;
}
