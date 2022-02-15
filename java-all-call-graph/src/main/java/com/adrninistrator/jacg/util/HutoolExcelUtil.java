package com.adrninistrator.jacg.util;

import cn.hutool.poi.excel.ExcelWriter;
import cn.hutool.poi.excel.StyleSet;
import org.apache.poi.ss.usermodel.*;

import static org.apache.poi.ss.usermodel.FillPatternType.SOLID_FOREGROUND;

/**
 * @Author: wqf
 * @Date: 2021/05/28
 * @Description: hutool 工具导出excel(非填充模板，手工画)
 */
public class HutoolExcelUtil {
    /**
     * YYYY/MM/dd 时间格式
     */
    private static final short LOCAL_DATE_FORMAT_SLASH = 14;

    /**
     * 方法描述: 全局基础样式设置
     * 默认 全局水平居中+垂直居中
     * 默认 自动换行
     * 默认单元格边框颜色为黑色，细线条
     * 默认背景颜色为白色
     *
     * @param writer writer
     * @param font   字体样式
     * @return cn.hutool.poi.excel.StyleSet
     * @author wqf
     * @date 2021/5/28 10:43
     */
    public static StyleSet setBaseGlobalStyle(ExcelWriter writer, Font font) {
        //全局样式设置
        StyleSet styleSet = writer.getStyleSet();
        //设置全局文本居中
        styleSet.setAlign(HorizontalAlignment.CENTER, VerticalAlignment.CENTER);
        //设置全局字体样式
        styleSet.setFont(font, true);
        //设置背景颜色 第二个参数表示是否将样式应用到头部
        styleSet.setBackgroundColor(IndexedColors.WHITE, true);
        //设置自动换行 当文本长于单元格宽度是否换行
        //styleSet.setWrapText();
        // 设置全局边框样式
        styleSet.setBorder(BorderStyle.THIN, IndexedColors.BLACK);
        return styleSet;
    }

    /**
     * 方法描述: 设置标题的基础样式
     *
     * @param styleSet            StyleSet
     * @param font                字体样式
     * @param horizontalAlignment 水平排列方式
     * @param verticalAlignment   垂直排列方式
     * @return org.apache.poi.ss.usermodel.CellStyle
     * @author wqf
     * @date 2021/5/28 10:16
     */
    public static CellStyle createHeadCellStyle(StyleSet styleSet, Font font,
                                                HorizontalAlignment horizontalAlignment,
                                                VerticalAlignment verticalAlignment) {
        CellStyle headCellStyle = styleSet.getHeadCellStyle();
        headCellStyle.setAlignment(horizontalAlignment);
        headCellStyle.setVerticalAlignment(verticalAlignment);
        headCellStyle.setFont(font);
        return headCellStyle;
    }

    /**
     * 方法描述: 设置基础字体样式字体 这里保留最基础的样式使用
     *
     * @param bold     是否粗体
     * @param fontName 字体名称
     * @param fontSize 字体大小
     * @return org.apache.poi.ss.usermodel.Font
     * @author wqf
     * @date 2021/5/19 15:58
     */
    public static Font createFont(ExcelWriter writer, boolean bold, boolean italic, String fontName, int fontSize) {
        Font font = writer.getWorkbook().createFont();
        //设置字体名称 宋体 / 微软雅黑 /等
        font.setFontName(fontName);
        //设置是否斜体
        font.setItalic(italic);
        //设置字体大小 以磅为单位
        font.setFontHeightInPoints((short) fontSize);
        //设置是否加粗
        font.setBold(bold);
        return font;
    }

    /**
     * 方法描述: 设置行或单元格基本样式
     *
     * @param writer              writer
     * @param font                字体样式
     * @param verticalAlignment   垂直居中
     * @param horizontalAlignment 水平居中
     * @return void
     * @author wqf
     * @date 2021/5/28 10:28
     */
    public static CellStyle createCellStyle(ExcelWriter writer, Font font, boolean wrapText,
                                            VerticalAlignment verticalAlignment,
                                            HorizontalAlignment horizontalAlignment) {
        CellStyle cellStyle = writer.getWorkbook().createCellStyle();
        cellStyle.setVerticalAlignment(verticalAlignment);
        cellStyle.setAlignment(horizontalAlignment);
        cellStyle.setWrapText(wrapText);
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 方法描述: 设置边框样式
     *
     * @param cellStyle 样式对象
     * @param bottom    下边框
     * @param left      左边框
     * @param right     右边框
     * @param top       上边框
     * @return void
     * @author wqf
     * @date 2021/5/28 10:23
     */
    public static void setBorderStyle(CellStyle cellStyle, BorderStyle bottom, BorderStyle left, BorderStyle right,
                                      BorderStyle top) {
        cellStyle.setBorderBottom(bottom);
        cellStyle.setBorderLeft(left);
        cellStyle.setBorderRight(right);
        cellStyle.setBorderTop(top);
    }


    /**
     * 方法描述: 自适应宽度(中文支持)
     *
     * @param sheet 页
     * @param size  因为for循环从0开始，size值为 列数-1
     * @return void
     * @author wqf
     * @date 2021/5/28 14:06
     */
    public static void setSizeColumn(Sheet sheet, int size) {
        for (int columnNum = 0; columnNum <= size; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }

                if (currentRow.getCell(columnNum) != null) {
                    Cell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }
    }

    /**
     * 单元格设置
     *
     * @param writer
     * @param borderTop           上边框
     * @param borderBottom        下边框
     * @param borderLeft          左边框
     * @param borderRight         右边框
     * @param isBold              是否加粗字体
     * @param horizontalAlignment 对齐方式
     * @return
     */
    public static CellStyle getCellStyle(ExcelWriter writer, BorderStyle borderTop, BorderStyle borderBottom, BorderStyle borderLeft, BorderStyle borderRight, boolean isBold, HorizontalAlignment horizontalAlignment) {
        CellStyle cellStyle = writer.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(borderTop);//上边框
        cellStyle.setBorderBottom(borderBottom);//下边框
        cellStyle.setBorderLeft(borderLeft);//左边框
        cellStyle.setBorderRight(borderRight);//右边框
        cellStyle.setAlignment(horizontalAlignment);//对齐方式
        Font font = writer.getWorkbook().createFont();
        font.setBold(isBold);
        font.setFontName("微软雅黑");//设置字体名称 宋体 / 微软雅黑 /等
        font.setFontHeightInPoints((short) 9); //字体大小
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 单元格设置
     *
     * @param writer
     * @param borderTop           上边框
     * @param borderBottom        下边框
     * @param borderLeft          左边框
     * @param borderRight         右边框
     * @param isBold              是否加粗字体
     * @param horizontalAlignment 对齐方式
     * @return
     */
    public static CellStyle getCellStyle(ExcelWriter writer, BorderStyle borderTop, BorderStyle borderBottom, BorderStyle borderLeft, BorderStyle borderRight, boolean isBold, HorizontalAlignment horizontalAlignment, Short indexedColors) {
        CellStyle cellStyle = writer.getWorkbook().createCellStyle();
        cellStyle.setBorderTop(borderTop);//上边框
        cellStyle.setBorderBottom(borderBottom);//下边框
        cellStyle.setBorderLeft(borderLeft);//左边框
        cellStyle.setBorderRight(borderRight);//右边框
        cellStyle.setAlignment(horizontalAlignment);//对齐方式
        Font font = writer.getWorkbook().createFont();
        font.setBold(isBold);
        font.setFontName("微软雅黑");//设置字体名称 宋体 / 微软雅黑 /等
        font.setFontHeightInPoints((short) 9); //字体大小
        cellStyle.setFont(font);

        cellStyle.setFillForegroundColor(indexedColors);
        cellStyle.setFillPattern(SOLID_FOREGROUND);
        return cellStyle;
    }

    /**
     * 这是单元格背景色
     *
     * @param writer
     * @param indexedColors
     * @return
     */
    public static CellStyle setFillBackgroundColor(ExcelWriter writer, Short indexedColors) {
        CellStyle cellStyle = writer.getWorkbook().createCellStyle();
        cellStyle.setFillForegroundColor(indexedColors);
        cellStyle.setFillPattern(SOLID_FOREGROUND);
        return cellStyle;
    }
}