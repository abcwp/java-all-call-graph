package com.maker;

import cn.hutool.core.collection.CollectionUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.crypto.SecureUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.adrninistrator.jacg.common.JACGConstants;
import com.adrninistrator.jacg.util.CommonUtil;
import com.adrninistrator.jacg.util.FileUtil;
import com.adrninistrator.jacg.util.HutoolExcelUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.CopyOnWriteArrayList;

import static com.adrninistrator.jacg.common.JACGConstants.EXT_XLSX;
import static com.adrninistrator.jacg.common.JACGConstants.FLAG_MINUS;
import static org.apache.poi.ss.usermodel.BorderStyle.NONE;
import static org.apache.poi.ss.usermodel.BorderStyle.THIN;

/**
 * @author wangpan
 * @date 2022/1/9
 */
public class CallGraphMaker {
    private static final Log logger = LogFactory.get();
    private Map<String, String> inputFilePaths = new HashMap<>(2);
    private String proInputFilePath;
    private String devInputFilePath;
    private String compareOutputFilePath;
    private Map<String, List<Map<String, String>>> datas = new HashMap<>(2);

    private static String[] headers1 = {"序号", "主键", "类名", "方法名", "方法被调用次数统计", "方法影响功能数统计", "备注"};
    private static String[] headers2 = {"主键", "类名", "方法名", "方法被调用次数统计", "方法影响功能数统计"};

    public static void main(String[] args) {
        try {
            long startTime = System.currentTimeMillis();
            new CallGraphMaker().run();
            long spendTime = System.currentTimeMillis() - startTime;
            logger.info("耗时: {} s", spendTime / 1000.0D);
        } catch (Exception e) {
            logger.error(e);
        }
    }

    public void run() throws Exception {
        //初始化
        init();

        //生成向上调用链可视化excel文件
        makeUpGraphView();
    }

    /**
     * 初始化
     */
    private void init() throws Exception {
        String configFilePath = JACGConstants.DIR_CONFIG + File.separator + JACGConstants.FILE_CONFIG;

        Reader reader = new InputStreamReader(new FileInputStream(FileUtil.findFile(configFilePath)), StandardCharsets.UTF_8);
        Properties properties = new Properties();
        properties.load(reader);

        String fullCalleeProFilePah = properties.getProperty(JACGConstants.COMBINE_FILE_NAME_4_CALLEE_PRO);
        if (StringUtils.isBlank(fullCalleeProFilePah)) {
            logger.error("配置文件中未指定参数：{} {}", configFilePath, JACGConstants.COMBINE_FILE_NAME_4_CALLEE_PRO);
        } else {
            logger.info("读取配置文件中指定参数：{} {}", configFilePath, JACGConstants.COMBINE_FILE_NAME_4_CALLEE_PRO);
            proInputFilePath = fullCalleeProFilePah;
            inputFilePaths.put("pro", proInputFilePath);
        }

        String fullCalleeDevFilePah = properties.getProperty(JACGConstants.COMBINE_FILE_NAME_4_CALLEE_DEV);
        if (StringUtils.isBlank(fullCalleeDevFilePah)) {
            logger.error("配置文件中未指定参数：{} {}", configFilePath, JACGConstants.COMBINE_FILE_NAME_4_CALLEE_DEV);
        } else {
            logger.info("读取配置文件中指定参数：{} {}", configFilePath, JACGConstants.COMBINE_FILE_NAME_4_CALLEE_DEV);
            devInputFilePath = fullCalleeDevFilePah;
            inputFilePaths.put("dev", devInputFilePath);
        }

        compareOutputFilePath = properties.getProperty(JACGConstants.COMBINE_FILE_NAME_4_CALLEE_COMPARE);
        logger.info("读取配置文件中指定参数：{} {}", configFilePath, JACGConstants.COMBINE_FILE_NAME_4_CALLEE_COMPARE);
    }

    /**
     * 生成向上调用链可视化excel分析文件
     */
    private void makeUpGraphView() throws Exception {
        if (CollectionUtil.isEmpty(inputFilePaths)) {
            logger.error("没有需要生成的文件");
            return;
        }
        for (Map.Entry<String, String> map : inputFilePaths.entrySet()) {
            makeUpGraphView(map.getKey(), map.getValue());
        }

        if (datas.size() == 2) {
            //生产、开发分支的数据都准备完毕，开始制作compare文件
            makeCompareView();
        }
    }

    /**
     * 根据指定文件，生成向上调用链可视化excel分析文件
     *
     * @param type          文件类型（pro、dev）
     * @param inputFilePath 文件路径
     */
    private void makeUpGraphView(String type, String inputFilePath) throws Exception {
        if (!FileUtil.isFileExists(inputFilePath)) {
            logger.error("输入文件不存在:{}", inputFilePath);
            return;
        }
        List<String> contentStrs = FileUtil.readFile2List(inputFilePath);
        if (CommonUtil.isCollectionEmpty(contentStrs)) {
            logger.error("输入文件内容为空:{}", inputFilePath);
            return;
        }
        //数据准备
        List<Map<String, String>> data = getData(contentStrs);
        logger.info("数据准备完成：{}", type);
        datas.put(type, data);

        //输出文件路径
        String dir = FileUtil.findFile(inputFilePath).getParent();
        String excelPath = dir + File.separator + type + FLAG_MINUS + DateUtil.currentSeconds() + EXT_XLSX;

        //制作excel
        makeExcel(excelPath, data);
    }

    /**
     * 数据解析
     *
     * @param contentStrs
     * @return
     */
    private synchronized List<Map<String, String>> getData(List<String> contentStrs) {
        List<Map<String, String>> dataList = new CopyOnWriteArrayList<>();
        String md5LineStr = "";//MD5加密
        String className = "";//类名
        String methodName = "";//方法名
        int methodCallCount = 0;//方法被调用次数统计
        int methodInfluenceFuncCount = 1;//方法影响功能数统计

        String lineStrTemp = "";
        int numTemp = 0;
        for (String lineStr : contentStrs) {
            String lineTrim = lineStr.trim();
            if (StringUtils.isBlank(lineTrim)) {
                continue;//剔除空行
            }
            // logger.info(lineTrim);

            if (!lineStr.startsWith(JACGConstants.FLAG_LEFT_PARENTHESES)) {
                if (StringUtils.isNotEmpty(lineStrTemp)) {
                    Map<String, String> mapTemp = new HashMap<>();
                    mapTemp.put("md5LineStr", md5LineStr);
                    mapTemp.put("className", className);
                    mapTemp.put("methodName", methodName);
                    mapTemp.put("methodCallCount", String.valueOf(methodCallCount));
                    mapTemp.put("methodInfluenceFuncCount", String.valueOf(methodInfluenceFuncCount));
                    dataList.add(mapTemp);
                }

                if (StringUtils.isEmpty(lineStrTemp) || !lineStrTemp.equals(lineStr)) {
                    md5LineStr = new String(SecureUtil.md5(lineStr));
                    String[] classAndMethod = lineStr.split(JACGConstants.FLAG_COLON);
                    className = classAndMethod[0];
                    methodName = classAndMethod[1];
                    methodCallCount = 0; //重置
                    methodInfluenceFuncCount = 1;//重置
                }
                lineStrTemp = lineStr;
                numTemp = 0;
            } else {
                Integer num = Integer.valueOf(StringUtils.substringBetween(lineStr, JACGConstants.FLAG_LEFT_PARENTHESES, JACGConstants.FLAG_RIGHT_PARENTHESES));
                if (Objects.isNull(num) || num == 0) {
                    continue;
                }
                if (num == 1) {
                    methodCallCount++;
                }
                if (num <= numTemp) {
                    methodInfluenceFuncCount++;
                }
                numTemp = num;
            }
        }
        //文件中最后一个方法数据添加到集合中
        Map<String, String> mapTemp = new HashMap<>();
        mapTemp.put("md5LineStr", md5LineStr);
        mapTemp.put("className", className);
        mapTemp.put("methodName", methodName);
        mapTemp.put("methodCallCount", String.valueOf(methodCallCount));
        mapTemp.put("methodInfluenceFuncCount", String.valueOf(methodInfluenceFuncCount));
        dataList.add(mapTemp);

        return dataList;
    }

    /**
     * 制作excel
     *
     * @param excelPath
     * @param data
     */
    private void makeExcel(String excelPath, List<Map<String, String>> data) throws Exception {
        OutputStream out = new FileOutputStream(excelPath);
        logger.info("准备制作文档：{}", excelPath);

        ExcelWriter writer = ExcelUtil.getWriter(true);
        Sheet sheet = writer.getSheet();
        sheet.setDisplayGridlines(false);//隐藏网格线
        sheet.setDefaultColumnWidth((short) 10);// 设置表格默认列宽度为20个字节

        Row row_2 = sheet.createRow(2);
        Cell cell_count = row_2.createCell(1);
        cell_count.setCellValue("方法数量：" + data.size());
        cell_count.setCellStyle(HutoolExcelUtil.getCellStyle(writer, NONE, NONE, NONE, NONE, false, HorizontalAlignment.LEFT));
        Cell cell_date = row_2.createCell(7);
        cell_date.setCellValue("统计日期：" + DateUtil.format(new Date(), "yyyy-MM-dd"));
        cell_date.setCellStyle(HutoolExcelUtil.getCellStyle(writer, NONE, NONE, NONE, NONE, false, HorizontalAlignment.RIGHT));

        Row row_3 = sheet.createRow(3);
        for (int i = 0; i < headers1.length; i++) {
            Cell cell = row_3.createCell(i + 1);
            cell.setCellStyle(HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, true, HorizontalAlignment.CENTER));
            cell.setCellValue(headers1[i]);
        }

        CellStyle cellStyle = HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, false, HorizontalAlignment.LEFT);//边框 左对齐
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 4);
            Map<String, String> map = data.get(i);

            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);
            Cell cell5 = row.createCell(5);
            Cell cell6 = row.createCell(6);
            Cell cell7 = row.createCell(7);
            cell1.setCellStyle(cellStyle);
            cell2.setCellStyle(cellStyle);
            cell3.setCellStyle(cellStyle);
            cell4.setCellStyle(cellStyle);
            cell5.setCellStyle(cellStyle);
            cell6.setCellStyle(cellStyle);
            cell7.setCellStyle(cellStyle);
            cell1.setCellValue(i + 1);
            cell2.setCellValue(map.get("md5LineStr"));
            cell3.setCellValue(map.get("className"));
            cell4.setCellValue(map.get("methodName"));
            cell5.setCellValue(map.get("methodCallCount"));
            cell6.setCellValue(map.get("methodInfluenceFuncCount"));
        }
        //writer.write(data);
        writer.flush(out);
        writer.close();
    }

    /**
     * 基于生产、开发分支的数据，制作compare文件
     */
    private void makeCompareView() throws Exception {
        logger.info("准备制作compare文件");
        if (StringUtils.isEmpty(compareOutputFilePath)) {
            logger.error("未指定输出文件存放路径:{}", compareOutputFilePath);
            return;
        }
        //输出文件路径
        String excelPath = compareOutputFilePath + File.separator + "compare" + FLAG_MINUS + DateUtil.currentSeconds() + EXT_XLSX;

        //制作excel
        makeCompareExcel(excelPath);
    }

    /**
     * 制作Compare excel
     *
     * @param excelPath
     */
    private void makeCompareExcel(String excelPath) throws Exception {
        Map<String, Map<String, String>> proData = dataPreproce(datas.get("pro"));
        Map<String, Map<String, String>> devData = dataPreproce(datas.get("dev"));

        OutputStream out = new FileOutputStream(excelPath);

        ExcelWriter writer = ExcelUtil.getWriter(true);
        Sheet sheet = writer.getSheet();
        sheet.setDisplayGridlines(false);//隐藏网格线
        sheet.setDefaultColumnWidth((short) 10);// 设置表格默认列宽度为20个字节

        makeCompareHead(writer, sheet);
        makeCompareBody(writer, sheet, proData, devData);

        writer.flush(out);
        writer.close();
    }

    /**
     * 制作compare文件的表头
     *
     * @param writer
     * @param sheet
     */
    private void makeCompareHead(ExcelWriter writer, Sheet sheet) {
        Cell cell_0_0 = sheet.createRow(0).createCell(0);//1行1列
        Cell cell_1_0 = sheet.createRow(1).createCell(0);//2行1列
        Cell cell_2_0 = sheet.createRow(2).createCell(0);//3行1列
        cell_0_0.setCellStyle(HutoolExcelUtil.setFillBackgroundColor(writer, IndexedColors.LIGHT_GREEN.index));
        cell_0_0.setCellValue("相较生产版本，新增方法");
        cell_1_0.setCellStyle(HutoolExcelUtil.setFillBackgroundColor(writer, IndexedColors.YELLOW.index));
        cell_1_0.setCellValue("相较生产版本，方法被调用次数/影响功能数量变动");
        cell_2_0.setCellStyle(HutoolExcelUtil.setFillBackgroundColor(writer, IndexedColors.RED.index));
        cell_2_0.setCellValue("相较生产版本，减少方法");

        Row row_4 = sheet.createRow(4);
        Cell cell_4_0 = row_4.createCell(0);
        Cell cell_4_6 = row_4.createCell(6);
        cell_4_0.setCellValue("PRO生产包");
        cell_4_6.setCellValue("DEV开发包");
        cell_4_0.setCellStyle(HutoolExcelUtil.setFillBackgroundColor(writer, IndexedColors.PALE_BLUE.index));
        cell_4_6.setCellStyle(HutoolExcelUtil.setFillBackgroundColor(writer, IndexedColors.PALE_BLUE.index));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 4));
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 6, 10));

        Row row_5 = sheet.createRow(5);
        for (int i = 0; i < headers2.length; i++) {
            Cell cell_pro = row_5.createCell(i);
            Cell cell_dev = row_5.createCell(i + 6);
            cell_pro.setCellStyle(HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, true, HorizontalAlignment.CENTER));
            cell_pro.setCellValue(headers2[i]);
            cell_dev.setCellStyle(HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, true, HorizontalAlignment.CENTER));
            cell_dev.setCellValue(headers2[i]);
        }
    }

    /**
     * 制作compare文件的body
     *
     * @param writer
     * @param sheet
     * @param proData
     * @param devData
     */
    private void makeCompareBody(ExcelWriter writer, Sheet sheet, Map<String, Map<String, String>> proData, Map<String, Map<String, String>> devData) {
        CellStyle cellStyle = HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, false, HorizontalAlignment.LEFT);//边框 左对齐
        CellStyle cellStyle_yellow = HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, false, HorizontalAlignment.LEFT, IndexedColors.YELLOW.index);//边框 左对齐 黄色填充
        CellStyle cellStyle_red = HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, false, HorizontalAlignment.LEFT, IndexedColors.RED.index);//边框 左对齐 红色填充
        int size = proData.size() > devData.size() ? proData.size() : devData.size();

        int rownum = 6;
        for (Map.Entry<String, Map<String, String>> proMap : proData.entrySet()) {
            boolean yellow = false;
            boolean red = false;
            String proMapKey = proMap.getKey();
            Map<String, String> proMapValue = proMap.getValue();
            String methodCallCount1 = proMapValue.get("methodCallCount");
            String methodInfluenceFuncCount1 = proMapValue.get("methodInfluenceFuncCount");

            Row row = sheet.createRow(rownum++);

            /*PRO数据*/
            Cell cell0 = row.createCell(0);
            Cell cell1 = row.createCell(1);
            Cell cell2 = row.createCell(2);
            Cell cell3 = row.createCell(3);
            Cell cell4 = row.createCell(4);
            cell0.setCellValue(proMapValue.get("md5LineStr"));
            cell1.setCellValue(proMapValue.get("className"));
            cell2.setCellValue(proMapValue.get("methodName"));
            cell3.setCellValue(proMapValue.get("methodCallCount"));
            cell4.setCellValue(proMapValue.get("methodInfluenceFuncCount"));
            /*PRO数据*/

            /*DEV数据*/
            Map<String, String> devMap = devData.get(proMapKey);
            Cell cell6 = row.createCell(6);
            Cell cell7 = row.createCell(7);
            Cell cell8 = row.createCell(8);
            Cell cell9 = row.createCell(9);
            Cell cell10 = row.createCell(10);
            if (CollectionUtil.isNotEmpty(devMap)) {
                cell6.setCellValue(devMap.get("md5LineStr"));
                cell7.setCellValue(devMap.get("className"));
                cell8.setCellValue(devMap.get("methodName"));
                cell9.setCellValue(devMap.get("methodCallCount"));
                cell10.setCellValue(devMap.get("methodInfluenceFuncCount"));

                String methodCallCount2 = devMap.get("methodCallCount");
                String methodInfluenceFuncCount2 = devMap.get("methodInfluenceFuncCount");

                //相较生产版本，方法被调用次数/影响功能数量变动
                if (!methodCallCount1.equals(methodCallCount2) || !methodInfluenceFuncCount1.equals(methodInfluenceFuncCount2)) {
                    yellow = true;
                }
            } else {
                //相较生产版本，减少方法
                red = true;
            }
            /*DEV数据*/
            CellStyle style = yellow ? cellStyle_yellow : red ? cellStyle_red : cellStyle;
            cell0.setCellStyle(style);
            cell1.setCellStyle(style);
            cell2.setCellStyle(style);
            cell3.setCellStyle(style);
            cell4.setCellStyle(style);
            cell6.setCellStyle(style);
            cell7.setCellStyle(style);
            cell8.setCellStyle(style);
            cell9.setCellStyle(style);
            cell10.setCellStyle(style);

            devData.remove(proMapKey);
        }

        CellStyle cellStyle_green = HutoolExcelUtil.getCellStyle(writer, THIN, THIN, THIN, THIN, false, HorizontalAlignment.LEFT, IndexedColors.LIGHT_GREEN.index);//边框 左对齐 红色填充
        for (Map.Entry<String, Map<String, String>> devMap : devData.entrySet()) {
            Map<String, String> devMapValue = devMap.getValue();
            Row row = sheet.createRow(rownum++);

            Cell cell6 = row.createCell(6);
            Cell cell7 = row.createCell(7);
            Cell cell8 = row.createCell(8);
            Cell cell9 = row.createCell(9);
            Cell cell10 = row.createCell(10);

            cell6.setCellValue(devMapValue.get("md5LineStr"));
            cell7.setCellValue(devMapValue.get("className"));
            cell8.setCellValue(devMapValue.get("methodName"));
            cell9.setCellValue(devMapValue.get("methodCallCount"));
            cell10.setCellValue(devMapValue.get("methodInfluenceFuncCount"));

            cell6.setCellStyle(cellStyle_green);
            cell7.setCellStyle(cellStyle_green);
            cell8.setCellStyle(cellStyle_green);
            cell9.setCellStyle(cellStyle_green);
            cell10.setCellStyle(cellStyle_green);
        }
    }

    /**
     * 数据预处理
     *
     * @param dataList
     * @return
     */
    private Map<String, Map<String, String>> dataPreproce(List<Map<String, String>> dataList) {
        Map<String, Map<String, String>> result = new HashMap<>();
        for (Map<String, String> map : dataList) {
            String md5LineStr = map.get("md5LineStr");
            if (!result.containsKey(md5LineStr)) {
                result.put(md5LineStr, map);
            }
        }
        return result;
    }
}
