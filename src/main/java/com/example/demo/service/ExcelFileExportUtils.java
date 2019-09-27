package com.example.demo.service;

import com.example.demo.entity.ExcelTitle;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFRegionUtil;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class ExcelFileExportUtils {
    private static Logger logger = LoggerFactory.getLogger(ExcelFileExportUtils.class);

    /**
     * 导出excel
     * @param map 标题行，存放标题的key和汉字值   例如：[id=ID,name=姓名,age=年龄]
     * @param list  内容，用于存放到excel内容区,存放的key要与标题中的key一致   例如：[[id=1,name=张三,age=20],[id=2,name=李四,age=30]]
     * @param filename  excel文件名
     * @param showOrderNumber 是否添加序号列
     * @param response
     */
    public static void exportExcelFile(LinkedHashMap<String,ExcelTitle> map, List<Map<String,Object>> list, String filename, boolean showOrderNumber, HttpServletResponse response){
        HSSFWorkbook workbook = new HSSFWorkbook();

        int theListSize = list.size(); //剩余处理数据量
        int sheetNum = 0;
        int sheetDateNum = 10000;//每个sheet页数据量
        logger.info("列表内容的大小："+list.size());
        do{

            HSSFSheet sheet = workbook.createSheet("sheet" + sheetNum);


            int titleStartRow = 0;
            int titleSize = map.size();//标题列数
            //第一个sheet页加大标题和时间
            if (sheetNum == 0){
                titleStartRow = 2;
                int bigTitleSize = titleSize;
                if (!showOrderNumber){
                    bigTitleSize = titleSize - 1;
                }
                //大标题行和时间行
                HSSFRow bigTitleRow = sheet.createRow(0);
                HSSFRow timeRow = sheet.createRow(1);
                //设置大标题样式
                HSSFCellStyle bigTitleCellStyle = workbook.createCellStyle();
                HSSFFont bigTitleFont = workbook.createFont();
                HSSFCellStyle timeCellStyle = workbook.createCellStyle();
                HSSFFont timeFont = workbook.createFont();
                //标题行样式
                bigTitleFont.setFontHeightInPoints((short) 20); //字体高度
                bigTitleFont.setBold(true);
                bigTitleFont.setFontName("宋体"); //字体
                bigTitleCellStyle.setFont(bigTitleFont);
                bigTitleCellStyle.setAlignment(HorizontalAlignment.CENTER);; //水平布局：居中
                HSSFCell bigTitleCell = bigTitleRow.createCell(0);
                bigTitleCell.setCellStyle(bigTitleCellStyle);
                bigTitleCell.setCellValue(filename);
                //时间行样式
                timeFont.setFontHeightInPoints((short) 12); //字体高度
                timeFont.setBold(false);
                timeFont.setFontName("宋体"); //字体
                timeCellStyle.setFont(timeFont);
                timeCellStyle.setAlignment(HorizontalAlignment.LEFT);; //水平布局：居中
                HSSFCell cell = timeRow.createCell(0);
                cell.setCellStyle(timeCellStyle);
                Date time = new Date();
                SimpleDateFormat tf = new SimpleDateFormat("yyyy年MM月dd日");
                cell.setCellValue("导出时间：" + tf.format(new Date()));

                //合并单元格
                CellRangeAddress bigTitle = new CellRangeAddress(0,0,0,bigTitleSize);
                CellRangeAddress timeTitle = new CellRangeAddress(1,1,0,bigTitleSize);
                //设置合并单元格边框
                HSSFRegionUtil.setBorderTop(1,bigTitle,sheet,workbook);
                HSSFRegionUtil.setBorderLeft(1,bigTitle,sheet,workbook);
                HSSFRegionUtil.setBorderRight(1,bigTitle,sheet,workbook);
                HSSFRegionUtil.setBorderLeft(1,timeTitle,sheet,workbook);
                HSSFRegionUtil.setBorderRight(1,timeTitle,sheet,workbook);
                HSSFRegionUtil.setBorderBottom(1,timeTitle,sheet,workbook);

                sheet.addMergedRegion(bigTitle);
                sheet.addMergedRegion(timeTitle);
            }

            //创建标题行
            HSSFRow row0 = sheet.createRow(titleStartRow);
            //设置标题样式
            HSSFCellStyle cellStyle = workbook.createCellStyle();
            HSSFFont font = workbook.createFont();
            //标题行样式
            font.setFontHeightInPoints((short) 16); //字体高度
            font.setBold(true);
            font.setFontName("宋体"); //字体
            cellStyle.setFont(font);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);; //水平布局：居中
            cellStyle.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex()); //标题背景色
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//填充图案的样式
            cellStyle.setBorderTop(BorderStyle.THIN);//上边框
            cellStyle.setBorderBottom(BorderStyle.THIN);//下边框
            cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
            cellStyle.setBorderRight(BorderStyle.THIN);//右边框

            int index = 0;
            if (showOrderNumber){
                index = 1;
                //添加序号列
                HSSFCell cell = row0.createCell(0);
                cell.setCellStyle(cellStyle);
                cell.setCellValue("序号");
            }
            for (String title:map.keySet()) {
                //创建标题列
                logger.info("title:"+title);
                HSSFCell cell = row0.createCell(index);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(map.get(title).getName());
                index++;
            }


            //内容行样式
            HSSFFont font1 = workbook.createFont();
            //内容区的样式
            font1.setFontHeightInPoints((short) 14); //字体高度
            font1.setFontName("宋体"); //字体

            HSSFCellStyle cellStyle1 = workbook.createCellStyle();//有背景色样式
            cellStyle1.setFont(font1);
            cellStyle1.setBorderTop(BorderStyle.THIN);//上边框
            cellStyle1.setBorderBottom(BorderStyle.THIN);//下边框
            cellStyle1.setBorderLeft(BorderStyle.THIN);//左边框
            cellStyle1.setBorderRight(BorderStyle.THIN);//右边框
            cellStyle1.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex()); //标题背景色
            cellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);//填充图案的样式

            HSSFCellStyle cellStyle2 = workbook.createCellStyle();//无背景色样式
            cellStyle2.setFont(font1);
            cellStyle2.setBorderTop(BorderStyle.THIN);//上边框
            cellStyle2.setBorderBottom(BorderStyle.THIN);//下边框
            cellStyle2.setBorderLeft(BorderStyle.THIN);//左边框
            cellStyle2.setBorderRight(BorderStyle.THIN);//右边框

            boolean groColorFlag = true;

            int nowSize = list.size()-sheetDateNum*sheetNum;
            logger.info("剩余数据量："+nowSize);
            if (nowSize > sheetDateNum){
                nowSize = sheetDateNum;
            }
            for (int i = 0; i < nowSize; i++) {
                //创建内容行
                HSSFRow row = sheet.createRow(i+1+titleStartRow);
                Map<String, Object> mapContent = list.get(i+sheetDateNum*sheetNum);
                if(map.size()!=mapContent.size()) {
                    logger.info("标题的列数："+map.size());
                    logger.info("内容的列数："+mapContent.size());
                    throw new RuntimeException("标题列数与内容列数不同！");
                }

                if (groColorFlag){
                    groColorFlag = false;
                }else{
                    groColorFlag = true;
                }
                int index1 = 0;
                if (showOrderNumber){
                    index1 = 1;
                    //序号
                    HSSFCell cell = row.createCell(0);
                    if (groColorFlag){
                        cell.setCellStyle(cellStyle1);
                    }else{
                        cell.setCellStyle(cellStyle2);
                    }
                    cell.setCellValue(sheetNum*sheetDateNum + i + 1);

                }
                for (String title:map.keySet()) {
                    Object value=null;
                    for(Map.Entry<String,Object> entrySet : mapContent.entrySet()) {
                        if(entrySet.getKey().equals(title)) {
                            //获取内容列的值
                            value = entrySet.getValue();
                        }
                    }
                    //创建内容区
                    HSSFCell cell = row.createCell(index1);
                    if (groColorFlag){
                        cell.setCellStyle(cellStyle1);
                    }else{
                        cell.setCellStyle(cellStyle2);
                    }

                    cell.setCellValue(value==null?"":String.valueOf(value));
                    index1++;
                }
            }

            //手动设置列宽行高
            /*sheet.setDefaultColumnWidth(20);
              sheet.setDefaultRowHeight((short)20);*/
            //设置自动调整宽度
            for(int i=0;i<titleSize;i++) {
                sheet.autoSizeColumn(i);
            }
            if (showOrderNumber){
                sheet.autoSizeColumn(titleSize);
            }

            theListSize -= sheetDateNum;
            sheetNum++;
        }while (theListSize>0);
        logger.info("数据处理完成");
        //流输出
        loadExportData(response,workbook,filename);
    }


    /**
     * 保存文件
     * @param response
     * @param workbook
     * @param filename
     */
    private static void loadExportData(HttpServletResponse response, HSSFWorkbook workbook, String filename) {
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            workbook.write(os);
            byte[] content = os.toByteArray();
            InputStream is = new ByteArrayInputStream(content);
            response.reset();
            response.setHeader("Pragma", "No-cache");//设置响应头信息，告诉浏览器不要缓存此内容
            response.setHeader("Cache-Control", "no-cache");
            response.setDateHeader("Expire", 0);
            response.setContentType("application/vnd.ms-excel;charset=UTF-8");
            //response.setHeader("Content-Disposition", "attachment;filename="+ new String((filename+".xls").getBytes(), "utf-8"));
            response.setHeader("Content-Disposition", "attachment;filename=" + new String(filename.getBytes(), "iso8859-1") + ".xls" );
            ServletOutputStream out = response.getOutputStream();
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[1024];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (Exception e) {
            logger.info("=====export exception=====", e);
        }finally {
            try {
                if(bis != null)
                    bis.close();
                if(bos != null)
                    bos.close();
            } catch (IOException e) {
                logger.info("=====close flow exception=====", e);
            }
        }
    }
}
