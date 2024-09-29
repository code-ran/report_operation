package com.ransibi.service.impl;

import com.alibaba.fastjson.JSONObject;
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

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;


@Service
public class UserServiceImpl implements IUserService {

    @Autowired
    UserMapper userMapper;

    private final static Map<String, String> VALUE_MAP = new HashMap<String, String>() {{
        put("0", "退");
        put("1", "投");
        put("-1", "--");
    }};

    private SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

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
//            workbook.getSheet("sheet名称");
        Sheet sheet = workbook.getSheetAt(0);
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
            Integer salary = ((Double) row.getCell(4).getNumericCellValue()).intValue();
//            Date hireDate = simpleDateFormat.parse(row.getCell(5).getStringCellValue());
            String hireDate = row.getCell(5).getStringCellValue();
//            Date birthDay = simpleDateFormat.parse(row.getCell(6).getStringCellValue());
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

        // 3、处理固定的标题  编号 姓名  手机号 入职日期 现住址
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        Row titleRow = sheet.createRow(0);
        Cell cell = null;
        for (int i = 0; i < 5; i++) {
            cell = titleRow.createCell(i);
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
    public void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("有样式的数据");

        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 8 * 256);
        sheet.setColumnWidth(2, 10 * 256);
        sheet.setColumnWidth(3, 10 * 256);
        sheet.setColumnWidth(4, 30 * 256);

        //需求：1、边框线：全边框  2、行高：42   3、合并单元格：第1行的第1个单元格到第5个单元格 4、对齐方式：水平垂直都要居中 5、字体：黑体18号字
        CellStyle bigTitleRowCellStyle = workbook.createCellStyle();
        bigTitleRowCellStyle.setBorderBottom(BorderStyle.THIN); //下边框  BorderStyle.THIN 细线
        bigTitleRowCellStyle.setBorderLeft(BorderStyle.THIN);  //左边框
        bigTitleRowCellStyle.setBorderRight(BorderStyle.THIN);  //右边框
        bigTitleRowCellStyle.setBorderTop(BorderStyle.THIN);  //上边框
        //对齐方式： 水平对齐  垂直对齐
        bigTitleRowCellStyle.setAlignment(HorizontalAlignment.CENTER); //水平居中对齐
        bigTitleRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中对齐
        //创建字体
        Font font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 18);
        //把字体放入到样式中
        bigTitleRowCellStyle.setFont(font);

        Row bigTitleRow = sheet.createRow(0);
        bigTitleRow.setHeightInPoints(42); //设置行高
        for (int i = 0; i < 5; i++) {
            Cell cell = bigTitleRow.createCell(i);
            cell.setCellStyle(bigTitleRowCellStyle);
        }
        //合并单元格: int firstRow 起始行, int lastRow 结束行, int firstCol 开始列, int lastCol 结束列
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        //向单元格中放入一句话
        sheet.getRow(0).getCell(0).setCellValue("用户信息数据");
        //小标题的样式
        CellStyle littleTitleRowCellStyle = workbook.createCellStyle();
        //样式的克隆
        littleTitleRowCellStyle.cloneStyleFrom(bigTitleRowCellStyle);
        //创建字体  宋体12号字加粗
        Font littleFont = workbook.createFont();
        littleFont.setFontName("宋体");
        littleFont.setFontHeightInPoints((short) 12);
        littleFont.setBold(true);
        //把字体放入到样式中
        littleTitleRowCellStyle.setFont(littleFont);
        //内容的样式
        CellStyle contentRowCellStyle = workbook.createCellStyle();
        //样式的克隆
        contentRowCellStyle.cloneStyleFrom(littleTitleRowCellStyle);
        contentRowCellStyle.setAlignment(HorizontalAlignment.LEFT);
        //创建字体  宋体12号字加粗
        Font contentFont = workbook.createFont();
        contentFont.setFontName("宋体");
        contentFont.setFontHeightInPoints((short) 11);
        contentFont.setBold(false);
        //把字体放入到样式中
        contentRowCellStyle.setFont(contentFont);
        //编号	姓名	手机号	入职日期	现住址
        Row titleRow = sheet.createRow(1);
        titleRow.setHeightInPoints(31.5F);
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        for (int i = 0; i < 5; i++) {
            Cell cell = titleRow.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(littleTitleRowCellStyle);
        }

        List<User> userList = userMapper.selectUserInfo();
        int rowIndex = 2;
        Row row = null;
        Cell cell = null;
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

        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());

    }

    @Override
    public void exportTest(HttpServletResponse response) throws Exception {
        List<ReCloseBaseBean> list = mockData();
        Workbook workbook = new XSSFWorkbook();
        //2、创建全新的工作表
        Sheet sheet = workbook.createSheet("测试sheet");
        // 设置表头样式
        CellStyle headerCellStyle = createHeadStyle(workbook);
        // 填充表头数据
        List<List<String>> headerData = new ArrayList<List<String>>();
        // 制造表头的数据
        headerData.add(Arrays.asList("地区,厂站,装置,结果,软压板,软压板,开关量,开关量,时间".split(",")));
        headerData.add(Arrays.asList("地区,厂站,装置,结果,名称,值,名称,值,时间".split(",")));
        for (int i = 0; i < headerData.size(); i++) {
            // 创建表头行
            final Row headerRow = sheet.createRow(i);
            for (int j = 0; j < headerData.get(i).size(); j++) {
                Cell cell = headerRow.createCell(j);
                cell.setCellValue(headerData.get(i).get(j).toString());
                cell.setCellStyle(headerCellStyle);
                sheet.setColumnWidth(j, 5000); //设置列宽度
            }
        }
        //表头垂直合并
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 3, 3));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 8, 8));
        //表头水平合并
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 5));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 6, 7));
        //表头已经占用的行数
        int tbHeadUseNo = headerData.size();
        // 设置表数据样式
        CellStyle tbStyle = createDataStyle(workbook);
        for (int rowIndex = 0; rowIndex < list.size(); rowIndex++) {
            ReCloseBaseBean bean = list.get(rowIndex);
            String sfInfo = bean.getSfInfo();
            String diInfo = bean.getDiInfo();
            String[] sfArray = new String[0];
            String[] diArray = new String[0];
            if (StringUtils.isNotEmpty(sfInfo) && !sfInfo.equals("NULL")) {
                sfArray = sfInfo.split(";");
            }
            if (StringUtils.isNotEmpty(diInfo) && !diInfo.equals("NULL")) {
                diArray = diInfo.split(";");
            }
            int maxLength = getMaxLength(sfArray, diArray);
            if (maxLength <= 1) {
                for (String s : sfArray) {
                    String[] item = s.split(",");
                    //名称
                    bean.setSfName(item[0]);
                    //值
                    bean.setSfValue(VALUE_MAP.get(item[1]));
                }
                for (String s : diArray) {
                    String[] item = s.split(",");
                    //名称
                    bean.setDiName(item[0]);
                    //值
                    bean.setDiValue(VALUE_MAP.get(item[1]));
                }
                if (sfArray.length > 0 && diArray.length == 0) {
                    //名称
                    bean.setDiName("--");
                    //值
                    bean.setDiValue("--");
                }
                if (diArray.length > 0 && sfArray.length == 0) {
                    //名称
                    bean.setSfName("--");
                    //值
                    bean.setSfValue("--");
                }
                //都为NULL的情况或者(sf/di中仅存在一组的情况)
                //直接生成一行,sf、di都为null
                Row row = sheet.createRow(rowIndex + tbHeadUseNo);
                //第一列
                Cell cell = row.createCell(0);
                cell.setCellValue(bean.getAreaName());
                cell.setCellStyle(tbStyle);
                //第二列
                cell = row.createCell(1);
                cell.setCellValue(bean.getStnName());
                cell.setCellStyle(tbStyle);
                //第三列
                cell = row.createCell(2);
                cell.setCellValue(bean.getPtName());
                cell.setCellStyle(tbStyle);
                //第四列
                cell = row.createCell(3);
                cell.setCellValue(bean.getIsAlarmName());
                cell.setCellStyle(tbStyle);
                //第五列
                cell = row.createCell(4);
                cell.setCellValue(bean.getSfName());
                cell.setCellStyle(tbStyle);
                //第六列
                cell = row.createCell(5);
                cell.setCellValue(bean.getSfValue());
                cell.setCellStyle(tbStyle);
                //第七列
                cell = row.createCell(6);
                cell.setCellValue(bean.getDiName());
                cell.setCellStyle(tbStyle);
                //第八列
                cell = row.createCell(7);
                cell.setCellValue(bean.getDiValue());
                cell.setCellStyle(tbStyle);
                //第九列
                cell = row.createCell(8);
                cell.setCellValue(bean.getChkTimeFormat());
                cell.setCellStyle(tbStyle);
            }
            if (maxLength == 2) {
                //sf/di中存在两组信息
                for (int i = 0; i < maxLength; i++) {
                    String[] item = sfArray[i].split(",");
                    String sfName = item[0];
                    String sfValue = VALUE_MAP.get(item[1]);
                    bean.setSfName(StringUtils.isNotEmpty(sfName) ? sfName : "--");
                    bean.setSfValue(StringUtils.isNotEmpty(sfValue) ? sfValue : "--");

                    String[] item1 = diArray[i].split(",");
                    String diName = item1[0];
                    String diValue = VALUE_MAP.get(item1[1]);
                    bean.setDiName(StringUtils.isNotEmpty(diName) ? diName : "--");
                    bean.setDiValue(StringUtils.isNotEmpty(diValue) ? diValue : "--");
                    // TODO: 2024/9/29 需要创建两行
                    Row row = sheet.createRow(rowIndex + tbHeadUseNo + i);
                    //第一列
                    Cell cell = row.createCell(0);
                    cell.setCellValue(bean.getAreaName());
                    cell.setCellStyle(tbStyle);
                    //第二列
                    cell = row.createCell(1);
                    cell.setCellValue(bean.getStnName());
                    cell.setCellStyle(tbStyle);
                    //第三列
                    cell = row.createCell(2);
                    cell.setCellValue(bean.getPtName());
                    cell.setCellStyle(tbStyle);
                    //第四列
                    cell = row.createCell(3);
                    cell.setCellValue(bean.getIsAlarmName());
                    cell.setCellStyle(tbStyle);
                    //第五列
                    cell = row.createCell(4);
                    cell.setCellValue(bean.getSfName());
                    cell.setCellStyle(tbStyle);
                    //第六列
                    cell = row.createCell(5);
                    cell.setCellValue(bean.getSfValue());
                    cell.setCellStyle(tbStyle);
                    //第七列
                    cell = row.createCell(6);
                    cell.setCellValue(bean.getDiName());
                    cell.setCellStyle(tbStyle);
                    //第八列
                    cell = row.createCell(7);
                    cell.setCellValue(bean.getDiValue());
                    cell.setCellStyle(tbStyle);
                    //第九列
                    cell = row.createCell(8);
                    cell.setCellValue(bean.getChkTimeFormat());
                    cell.setCellStyle(tbStyle);
                }
                int lastRowNum = sheet.getLastRowNum();
                // TODO: 2024/9/29 合并上面创建的两行
                //地区合并
                sheet.addMergedRegion(new CellRangeAddress(lastRowNum - 1, lastRowNum, 0, 0));
                //厂站合并
                sheet.addMergedRegion(new CellRangeAddress(lastRowNum - 1, lastRowNum, 1, 1));
                //装置合并
                sheet.addMergedRegion(new CellRangeAddress(lastRowNum - 1, lastRowNum, 2, 2));
                //结果合并
                sheet.addMergedRegion(new CellRangeAddress(lastRowNum - 1, lastRowNum, 3, 3));
                //时间合并
                sheet.addMergedRegion(new CellRangeAddress(lastRowNum - 1, lastRowNum, 8, 8));
            }
        }
        //  一个流两个头
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }


    private List<ReCloseBaseBean> mockData() {
        List<ReCloseBaseBean> reCloseBaseBeanList = new ArrayList<>();
        ReCloseBaseBean obj = new ReCloseBaseBean();
        //地区
        obj.setAreaName("白银公司");
        //厂站
        obj.setStnName("220kV沙河变");
        //装置
        obj.setPtName("110kV母联断路器1100保护PSL-633U");
        //结果
        obj.setIsAlarmName("正常");
        obj.setSfInfo("NULL");
        obj.setDiInfo("NULL");
        //时间
        obj.setChkTime(new Date());
        reCloseBaseBeanList.add(obj);
        ReCloseBaseBean obj1 = new ReCloseBaseBean();
        //地区
        obj1.setAreaName("白银公司");
        //厂站
        obj1.setStnName("220kV沙河变");
        //装置
        obj1.setPtName("220kV母联断路器2200第二套保护PCS-923A-G");
        //结果
        obj1.setIsAlarmName("异常");
        obj1.setSfInfo("NULL");
        obj1.setDiInfo("充电过流保护软压板,0");
        //时间
        obj1.setChkTime(new Date());
        reCloseBaseBeanList.add(obj1);
        ReCloseBaseBean obj2 = new ReCloseBaseBean();
        //地区
        obj2.setAreaName("白银公司");
        //厂站
        obj2.setStnName("330kV东台变");
        //装置
        obj2.setPtName("330kV断路器3321保护NSR-321A-G");
        //结果
        obj2.setIsAlarmName("异常");
        obj2.setSfInfo("充电过流保护软压板,1");
        obj2.setDiInfo("充电过流保护硬压板,0");
        //时间
        obj2.setChkTime(new Date());
        reCloseBaseBeanList.add(obj2);
        ReCloseBaseBean obj3 = new ReCloseBaseBean();
        //地区
        obj3.setAreaName("白银公司");
        //厂站
        obj3.setStnName("330kV东台变");
        //装置
        obj3.setPtName("330kV断路器3322保护NSR-322A-G");
        //结果
        obj3.setIsAlarmName("正常");
        obj3.setSfInfo("充电过流保护软压板,1;充电过流保护硬压板,0");
        obj3.setDiInfo("充电过流保护硬压板,0;充电过流保护软压板,-1");
        //时间
        obj3.setChkTime(new Date());
        reCloseBaseBeanList.add(obj3);
        return reCloseBaseBeanList;
    }

    /**
     * 创建表头样式.
     *
     * @param
     * @param
     */
    private static CellStyle createHeadStyle(final Workbook workbook) {
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        headerCellStyle.setBorderBottom(BorderStyle.THIN);// 下边框
        headerCellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        headerCellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        headerCellStyle.setBorderRight(BorderStyle.THIN);// 右边框
//        headerCellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.index);//蓝色背景色
//        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//全填充模式
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());// 背景颜色
        Font font = workbook.createFont();
//        font.setColor(IndexedColors.WHITE.getIndex());//设置字体颜色
//        font.setBold(true);
//        headerCellStyle.setFont(font);//表头字体加粗
        font.setBold(true); // 字体加粗
        font.setFontName("黑体"); // 设置字体类型
        font.setFontHeightInPoints((short) 15); // 设置字体大小
        headerCellStyle.setFont(font); // 为标题样式设置字体样式
        return headerCellStyle;
    }

    /**
     * 创建数据样式.
     *
     * @param
     * @param
     */
    private static CellStyle createDataStyle(final Workbook workbook) {
        CellStyle tbStyle = workbook.createCellStyle();
        tbStyle.setAlignment(HorizontalAlignment.CENTER);//水平居中
        tbStyle.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        tbStyle.setWrapText(true);// 设置自动换行
        tbStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        tbStyle.setBorderLeft(BorderStyle.THIN); // 左边框
        tbStyle.setBorderRight(BorderStyle.THIN); // 右边框
        tbStyle.setBorderTop(BorderStyle.THIN); // 上边框
        Font tbfont = workbook.createFont();
        tbfont.setColor((short) 8);
        tbfont.setFontHeightInPoints((short) 12);
        tbStyle.setFont(tbfont);
        return tbStyle;
    }

    private static int getMaxLength(String[] sfArray, String[] diArray) {
        int max = 0;
        if (diArray.length >= sfArray.length) {
            max = diArray.length;
        } else if (sfArray.length >= diArray.length) {
            max = sfArray.length;
        }
        return max;
    }
}
