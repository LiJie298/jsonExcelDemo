package com.lijie.util;

import org.apache.logging.log4j.util.Strings;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author lijie7
 * @date 2017/12/4
 * @Description
 * @modified By
 */
public class ExcelFileUtil {

    /**
     * 获取cell中值 转换成String类型
     *
     * @param cell
     * @return
     */
    public static Object getCellValue(Cell cell) {
        Object cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case XSSFCell.CELL_TYPE_NUMERIC: {
                    short format = cell.getCellStyle().getDataFormat();
                    //excel中的时间格式
                    if (format == 14 || format == 31 || format == 57 || format == 58) {
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM");
                        double value = cell.getNumericCellValue();
                        Date date = DateUtil.getJavaDate(value);
                        cellvalue = sdf.format(date);
                    }else if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        // 判断当前的cell是否为Date
                        // 先注释日期类型的转换，在实际测试中发现HSSFDateUtil.isCellDateFormatted(cell)只识别2014/02/02这种格式。
                        // 如果是Date类型则，取得该Cell的Date值
                        // 对2014-02-02格式识别不出是日期格式
                        Date date = cell.getDateCellValue();
                        DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = formater.format(date);
                    } else {
                        // 如果是纯数字 取得当前Cell的数值
                        cellvalue = cell.getNumericCellValue();
                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case XSSFCell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getStringCellValue().replaceAll("'", "''");
                    if ("-".equals(cellvalue)) {
                        cellvalue = " ";
                    }
                    break;
                case XSSFCell.CELL_TYPE_BLANK:
                    cellvalue = null;
                    break;
                // 默认的Cell值
                default:
                    cellvalue = null;
            }
        } else {
            cellvalue = null;
        }
        return cellvalue;
    }


    /**
     * 判断单元格是否是合并单元格
     *
     * @param sheet
     * @param row
     * @param column
     * @return
     */
    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取合并单元格数值
     *
     * @param sheet
     * @param cell
     * @return
     */
    public static Object getMergedRegionValue(Sheet sheet, Cell cell) {
        int row = cell.getRowIndex(), column = cell.getColumnIndex();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell);
                }
            }
        }
        return null;
    }



    /**
     * 设置表头风格
     *
     * @param style
     * @param font
     * @param index
     * @return
     */

    public static XSSFCellStyle getTitleCellStyle(XSSFCellStyle style, Font font, short index) {
        font.setColor(IndexedColors.WHITE.index);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        return getBasicStyle(style, index);
    }

    /**
     * 设置内容风格
     *
     * @param style
     * @param font
     * @param index 填充颜色
     * @return
     */
    public static XSSFCellStyle getContentCellStyle(XSSFCellStyle style, Font font, short index) {
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        return getBasicStyle(style, index);
    }

    //设置基本样式
    public static XSSFCellStyle getBasicStyle(XSSFCellStyle style, short index) {
        if (index == 0) {
            style.setFillForegroundColor(IndexedColors.WHITE.index);
            style.setFillBackgroundColor(IndexedColors.WHITE.index);
        } else {
            style.setFillForegroundColor(index);
            style.setFillBackgroundColor(index);
        }
        //背景色填充
        style.setFillPattern(XSSFCellStyle.BIG_SPOTS);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        return style;
    }

    public static XSSFWorkbook setAutoColumnSize(XSSFWorkbook workbook) throws Exception {
        return setAutoColumnSize(workbook, 0, 0);
    }

    public static XSSFWorkbook setAutoColumnSize(XSSFWorkbook workbook, int sheetIndex, int rowIndex) throws Exception {
        XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
        int maxColumn = sheet.getRow(rowIndex).getPhysicalNumberOfCells();
        for (int i = 0; i <= maxColumn; i++) {
            sheet.autoSizeColumn(i);
        }
        //获取当前列的宽度，然后对比本列的长度，取最大值
        for (int columnNum = 0; columnNum <= maxColumn; columnNum++) {
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
                    int length = currentCell.toString().getBytes("GBK").length;
                    if (columnWidth < length + 1) {
                        columnWidth = length + 1;
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256 > 65280 / 2 ? 65280 / 2 : columnWidth * 256);
        }
        return workbook;
    }

    /**
     * 设置数据单元格格式
     *
     * @param cell
     * @param a
     */
    public static void setCellFormatValue(XSSFCell cell, Object a, XSSFCellStyle stringStyle, XSSFCellStyle intStyle, XSSFCellStyle doubleStyle) {
        if (a == null || Strings.isBlank(a.toString())) {
            cell.setCellStyle(stringStyle);
            cell.setCellValue("");
            return;
        }
        Class a1 = a.getClass();
        if (a1 == String.class) {
            cell.setCellStyle(stringStyle);
            cell.setCellValue(a.toString());
            return;
        } else {
            if (a1 == Integer.class) {
                cell.setCellStyle(intStyle);
                cell.setCellValue(Integer.parseInt(a.toString()));
                return;
            } else if (a1 == Double.class) {
                cell.setCellStyle(doubleStyle);
                cell.setCellValue(Double.parseDouble(a.toString()));
                return;
            } else {
                cell.setCellStyle(stringStyle);
                cell.setCellValue(a.toString());
                return;
            }
        }
    }


}
