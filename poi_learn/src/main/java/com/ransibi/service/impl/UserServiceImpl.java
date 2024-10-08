package com.ransibi.service.impl;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.ransibi.dao.UserMapper;
import com.ransibi.pojo.ReCloseBaseBean;
import com.ransibi.pojo.User;
import com.ransibi.service.IUserService;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;


@Service
public class UserServiceImpl implements IUserService {

    @Autowired
    UserMapper userMapper;

    @Override
    public List<User> getUserInfo(Integer page, Integer pageSize) {
        PageHelper.startPage(page, pageSize);
        Page<User> userPage = (Page<User>) userMapper.selectUserInfo();
        return userPage.getResult();
    }

    @Override
    public String uploadExcelInfo(MultipartFile file) throws Exception {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        //获取第一个sheet
        //Sheet sheet = workbook.getSheet("sheet名称");
        Sheet sheet = workbook.getSheetAt(0);
        //获取最后一行索引
        int lastRowIndex = sheet.getLastRowNum();
        Row row = null;
        List<User> userList = new ArrayList<>();
        //从第二行开始读
        for (int i = 1; i <= lastRowIndex; i++) {
            //获取行数据
            row = sheet.getRow(i);
            String userName = row.getCell(0).getStringCellValue();
            String phone = null;
            try {
                phone = row.getCell(1).getStringCellValue();
            } catch (Exception e) {
                phone = row.getCell(1).getNumericCellValue() + "";
            }
            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            //工资是数值类型使用getNumericCellValue()方法获取
            Integer salary = ((Double) row.getCell(4).getNumericCellValue()).intValue();
            String hireDate = row.getCell(5).getStringCellValue();
            String birthDay = row.getCell(6).getStringCellValue();
            String address = row.getCell(7).getStringCellValue();
            User user = new User();
            user.setUserName(userName);
            user.setPhone(phone);
            user.setProvince(province);
            user.setCity(city);
            user.setSalary(salary);
            user.setHireDateFormat(hireDate);
            user.setBirthdayFormat(birthDay);
            user.setAddress(address);
            user.setDeptId(4);
            userList.add(user);
        }
        if (CollectionUtils.isNotEmpty(userList)) {
            userMapper.insertUser(userList);
        }
        return "导入成功";
    }

    @Override
    public void downLoadXlsxByPoi(HttpServletResponse response) throws Exception {
        /*导出用户数据基本思路：
        1、创建一个全新的工作薄
        2、创建全新的工作表
        3、处理固定的标题  编号 姓名  手机号 入职日期 现住址
        4、从第二行开始循环遍历 向单元格中放入数据*/
        // 1、创建一个全新的工作薄
        Workbook workbook = new XSSFWorkbook();
        //2、创建全新的工作表
        Sheet sheet = workbook.createSheet("用户数据");
        // 设置列宽
        // 1代表的是一个标准字母宽度的256分之一
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 8 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 15 * 256);
        sheet.setColumnWidth(4, 30 * 256);

        // 3、处理固定的标题: 编号、姓名 、手机号、入职日期、现住址
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        Row titleRow = sheet.createRow(0);
        Cell cell = null;
        for (int i = 0; i < 5; i++) {
            //创建单元格
            cell = titleRow.createCell(i);
            //填充数据
            cell.setCellValue(titles[i]);
        }
        // 4、从第二行开始循环遍历 向单元格中放入数据
        List<User> userList = userMapper.selectUserInfo();
        int rowIndex = 1;
        Row row = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            // 编号 姓名  手机号 入职日期 现住址
            cell = row.createCell(0);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellValue(user.getPhone());

            cell = row.createCell(3);
            cell.setCellValue(user.getHireDateFormat());

            cell = row.createCell(4);
            cell.setCellValue(user.getAddress());

            rowIndex++;
        }
        //  一个流两个头
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }

    //    使用POI导出用户列表数据--带样式
    @Override
    public void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("用户数据");
        //设置大标题
        Row bigTitleRow = sheet.createRow(0);
        //设置行高
        bigTitleRow.setHeightInPoints(42);
        CellStyle bigTitleRowCellStyle = bigTitleRowCellStyle(workbook);
        //第一行数据填充
        for (int i = 0; i < 5; i++) {
            Cell cell = bigTitleRow.createCell(i);
            cell.setCellStyle(bigTitleRowCellStyle);
        }
        //合并单元格: int firstRow 起始行, int lastRow 结束行, int firstCol 开始列, int lastCol 结束列
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        //向单元格中放入一句话
        sheet.getRow(0).getCell(0).setCellValue("用户信息数据统计");
        //设置小标题
        //编号	姓名	手机号	入职日期	现住址
        Row titleRow = sheet.createRow(1);
        titleRow.setHeightInPoints(31.5F);
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        CellStyle littleTitleRowCellStyle = littleTitleRowCellStyle(workbook);
        for (int i = 0; i < 5; i++) {
            Cell cell = titleRow.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(littleTitleRowCellStyle);
        }
        //数据填充
        List<User> userList = userMapper.selectUserInfo();
        int rowIndex = 2;
        Row row = null;
        Cell cell = null;
        CellStyle contentRowCellStyle = contentRowCellStyle(workbook);
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            cell = row.createCell(0);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getPhone());

            cell = row.createCell(3);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getHireDateFormat());

            cell = row.createCell(4);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getAddress());

            rowIndex++;
        }
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 12 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 20 * 256);
        sheet.setColumnWidth(4, 30 * 256);
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }

    @Override
    public void downLoadXlsxWithTemplate(HttpServletResponse response) throws Exception {
        //获取模板的路径
        String templatePath = Class.class.getClass().getResource("/").getPath() + "template/user_template.xlsx";
        //读取模板文件产生workbook对象,这个workbook是一个有内容的工作薄
        Workbook workbook = new XSSFWorkbook(new File(templatePath));
        //读取工作薄的第一个工作表，向工作表中放数据
        Sheet sheet = workbook.getSheetAt(0);
        //数据填充
        List<User> userList = userMapper.selectUserInfo();
        int rowIndex = 2;
        Row row = null;
        Cell cell = null;
        CellStyle contentRowCellStyle = contentRowCellStyle(workbook);
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            cell = row.createCell(0);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getPhone());

            cell = row.createCell(3);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getHireDateFormat());

            cell = row.createCell(4);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getAddress());

            rowIndex++;
        }
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 12 * 256);
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 20 * 256);
        sheet.setColumnWidth(4, 30 * 256);
        //导出的文件名称
        String filename = "用户列表数据.xlsx";
        //设置文件的打开方式和mime类型
        ServletOutputStream outputStream = response.getOutputStream();
        response.setHeader("Content-Disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(outputStream);
    }


    /**
     * 公共样式
     *
     * @param workbook
     * @return
     */
    private static CellStyle commonStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        //下边框  BorderStyle.THIN 细线
        cellStyle.setBorderBottom(BorderStyle.THIN);
        //左边框
        cellStyle.setBorderLeft(BorderStyle.THIN);
        //右边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        //上边框
        cellStyle.setBorderTop(BorderStyle.THIN);
        //对齐方式： 水平对齐  垂直对齐
        //水平居中对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中对齐
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return cellStyle;
    }

    /**
     * 大标题样式设置
     *
     * @param workbook
     * @return
     */
    private static CellStyle bigTitleRowCellStyle(Workbook workbook) {
        CellStyle cellStyle = commonStyle(workbook);
        //创建字体
        Font font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 18);
        //把字体放入到样式中
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 小标题样式设置
     *
     * @param workbook
     * @throws Exception
     */
    private static CellStyle littleTitleRowCellStyle(Workbook workbook) {
        CellStyle cellStyle = commonStyle(workbook);
        //创建字体  宋体12号字加粗
        Font littleFont = workbook.createFont();
        littleFont.setFontName("宋体");
        littleFont.setFontHeightInPoints((short) 12);
        littleFont.setBold(true);
        //把字体放入到样式中
        cellStyle.setFont(littleFont);
        return cellStyle;
    }

    /**
     * 内容样式设置
     *
     * @param workbook
     * @return
     */
    private static CellStyle contentRowCellStyle(Workbook workbook) {
        //内容的样式
        CellStyle cellStyle = workbook.createCellStyle();
        //样式的克隆
        cellStyle.cloneStyleFrom(littleTitleRowCellStyle(workbook));
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        //创建字体  宋体12号字加粗
        Font contentFont = workbook.createFont();
        contentFont.setFontName("宋体");
        contentFont.setFontHeightInPoints((short) 11);
        contentFont.setBold(false);
        //把字体放入到样式中
        cellStyle.setFont(contentFont);
        return cellStyle;
    }
}
