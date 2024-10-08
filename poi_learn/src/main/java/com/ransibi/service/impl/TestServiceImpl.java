package com.ransibi.service.impl;

import com.ransibi.dao.UserMapper;
import com.ransibi.pojo.ReCloseBaseBean;
import com.ransibi.service.ITestService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.util.*;

@Service
public class TestServiceImpl implements ITestService {

    @Autowired
    UserMapper userMapper;

    private final static Map<String, String> VALUE_MAP = new HashMap<String, String>() {{
        put("0", "退");
        put("1", "投");
        put("-1", "--");
    }};

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
            String[] sfArray = StringUtils.isNotEmpty(bean.getSfInfo()) && !bean.getSfInfo().equals("NULL")
                    ? bean.getSfInfo().split(";") : new String[0];
            String[] diArray = StringUtils.isNotEmpty(bean.getDiInfo()) && !bean.getDiInfo().equals("NULL")
                    ? bean.getDiInfo().split(";")
                    : new String[0];
            int maxLength = getMaxLength(sfArray, diArray);
            // TODO: 2024/10/8
            if (maxLength <= 1) {
                processSingleRow(sfArray, diArray, bean);
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
            } else if (maxLength == 2) {
                for (int i = 0; i < maxLength; i++) {
                    processMultipleRows(sfArray, diArray, bean, i);
                    int lastRowNumTemp = sheet.getLastRowNum();
                    Row row = sheet.createRow(lastRowNumTemp + 1);
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

    private void processSingleRow(String[] sfArray, String[] diArray, ReCloseBaseBean bean) {
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
    }

    private void processMultipleRows(String[] sfArray, String[] diArray, ReCloseBaseBean bean, int i) {
        if (sfArray.length >= diArray.length) {
            String[] item = sfArray[i].split(",");
            String sfName = item[0];
            String sfValue = VALUE_MAP.get(item[1]);
            bean.setSfName(StringUtils.isNotEmpty(sfName) ? sfName : "--");
            bean.setSfValue(StringUtils.isNotEmpty(sfValue) ? sfValue : "--");
            if (diArray.length == 2) {
                String[] item1 = diArray[i].split(",");
                String diName = item1[0];
                String diValue = VALUE_MAP.get(item1[1]);
                bean.setDiName(StringUtils.isNotEmpty(diName) ? diName : "--");
                bean.setDiValue(StringUtils.isNotEmpty(diValue) ? diValue : "--");
            }
            if (i == 0 && diArray.length == 1) {
                String[] item1 = diArray[0].split(",");
                String diName = item1[0];
                String diValue = VALUE_MAP.get(item1[1]);
                bean.setDiName(StringUtils.isNotEmpty(diName) ? diName : "--");
                bean.setDiValue(StringUtils.isNotEmpty(diValue) ? diValue : "--");
            }
            if ((i == 1 && diArray.length == 1) || (diArray.length == 0)) {
                bean.setDiName("--");
                bean.setDiValue("--");
            }
        } else {
            String[] item1 = diArray[i].split(",");
            String diName = item1[0];
            String diValue = VALUE_MAP.get(item1[1]);
            bean.setDiName(StringUtils.isNotEmpty(diName) ? diName : "--");
            bean.setDiValue(StringUtils.isNotEmpty(diValue) ? diValue : "--");
            if (sfArray.length == 2) {
                String[] item = sfArray[i].split(",");
                String sfName = item[0];
                String sfValue = VALUE_MAP.get(item[1]);
                bean.setSfName(StringUtils.isNotEmpty(sfName) ? sfName : "--");
                bean.setSfValue(StringUtils.isNotEmpty(sfValue) ? sfValue : "--");
            }
            if (i == 0 && sfArray.length == 1) {
                String[] item = sfArray[0].split(",");
                String sfName = item[0];
                String sfValue = VALUE_MAP.get(item[1]);
                bean.setSfName(StringUtils.isNotEmpty(sfName) ? sfName : "--");
                bean.setSfValue(StringUtils.isNotEmpty(sfValue) ? sfValue : "--");
            }
            if ((i == 1 && sfArray.length == 1) || (sfArray.length == 0)) {
                bean.setSfName("--");
                bean.setSfValue("--");
            }
        }
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


        ReCloseBaseBean obj5 = new ReCloseBaseBean();
        //地区
        obj5.setAreaName("白银公司");
        //厂站
        obj5.setStnName("330kV东台变");
        //装置
        obj5.setPtName("330kV断路器3324保护NSR-322A-G");
        //结果
        obj5.setIsAlarmName("异常");
        obj5.setSfInfo("充电过流保护软压板,1");
        obj5.setDiInfo("充电过流保护硬压板,0;充电过流保护软压板,-1");
        //时间
        obj5.setChkTime(new Date());
        reCloseBaseBeanList.add(obj5);

//        ReCloseBaseBean obj4 = new ReCloseBaseBean();
//        //地区
//        obj4.setAreaName("白银公司");
//        //厂站
//        obj4.setStnName("330kV东台变");
//        //装置
//        obj4.setPtName("330kV断路器3322保护NSR-333A-G");
//        //结果
//        obj4.setIsAlarmName("正常");
//        obj4.setSfInfo("NULL");
//        obj4.setDiInfo("充电过流保护硬压板,0;充电过流保护软压板,-1");
//        //时间
//        obj4.setChkTime(new Date());
//        reCloseBaseBeanList.add(obj4);
        return reCloseBaseBeanList;
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
}
