package com.ransibi.service;


import com.ransibi.pojo.User;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

public interface IUserService {
    /**
     * 获取列表数据
     * @param page
     * @param pageSize
     * @return
     */
    List<User> getUserInfo(Integer page, Integer pageSize);
    /**
     * 通过poi导入Excel数据
     * @param file
     * @return
     * @throws Exception
     */
    String uploadExcelInfo(MultipartFile file) throws Exception;
    /**
     * 通过poi导出Excel
     * @param response
     * @throws Exception
     */
    void downLoadXlsxByPoi(HttpServletResponse response) throws Exception;
    /**
     * 导出带样式的Excel
     * @param response
     * @throws Exception
     */
    void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws Exception;
    /**
     * 基于模版导出Excel
     * @param response
     * @throws Exception
     */
    void downLoadXlsxWithTemplate(HttpServletResponse response) throws Exception;

    /**
     * 通过模版导出详细数据Excel
     * @param id
     * @param response
     * @throws Exception
     */
    void downLoadFileInfo(Long id,HttpServletResponse response) throws Exception;
}
