package com.lijie.service;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.lijie.excelFile.ExcelFile;
import com.lijie.excelFile.ExcelSheet;
import com.lijie.excelFile.ExcelTitleTem;
import com.lijie.util.ExcelFileUtil;
import com.lijie.util.Log;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import static com.lijie.util.ExcelFileUtil.getContentCellStyle;
import static com.lijie.util.ExcelFileUtil.setCellFormatValue;


/**
 * @author lijie7
 * @date 2018/5/25
 * @Description
 * @modified By
 */
@Service
public class ExcelDowService {


    /**
     * 文件数据导出
     *
     * @param excelFile
     * @param jsonObject
     * @return
     * @throws Exception
     */
    public ByteArrayOutputStream exportFile(ExcelFile excelFile, JSONObject jsonObject) {
        ByteArrayOutputStream bos = null;
        try {
            if (!jsonObject.get("datas").getClass().getTypeName().equalsIgnoreCase("JSONObject")) {
                bos = getReportFileByJsonArray(excelFile, jsonObject.getJSONArray("datas"));
            } else {
                bos = getReportFileByJsonObject(excelFile, jsonObject.getJSONObject("datas"));
            }
        } catch (Exception e) {
            Log.error("文件生成错误", e.toString());
        }
        return bos;
    }

    public static void main(String[] args) {
        ExcelDowService service = new ExcelDowService();
        ExcelFile file = new ExcelFile("demo3");

        try {
            String jj = getdata();
            ByteArrayOutputStream byteArrayOutputStream = getReportFileByJsonArray(file, (JSONArray) JSONArray.parse(jj));
            File file1 = new File("D:\\key.xlsx");
            FileOutputStream fileOutputStream = new FileOutputStream(file1);
            fileOutputStream.write(byteArrayOutputStream.toByteArray());
            fileOutputStream.flush();
            fileOutputStream.close();
            byteArrayOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String getdata() {
        StringBuffer stringBuffer = new StringBuffer();
        try{

            File file  = new File("D:\\My_Code\\jsonExcelDemo\\nianhuizhongjiang.json");
            FileInputStream fileInputStream = new FileInputStream(file);
            InputStreamReader isr=new InputStreamReader(fileInputStream, "UTF-8");
            BufferedReader br = new BufferedReader(isr);
            //简写如下
            //BufferedReader br = new BufferedReader(new InputStreamReader(
            //        new FileInputStream("E:/phsftp/evdokey/evdokey_201103221556.txt"), "UTF-8"));
            String line="";
            String[] arrs=null;
            while ((line=br.readLine())!=null) {
                stringBuffer.append(line);
            }
            br.close();
            isr.close();
            fileInputStream.close();
        }catch (Exception er){
            er.printStackTrace();
        }
        return stringBuffer.toString();
    }

    /**
     * 报表导出
     *
     * @param jsonArray
     * @return
     * @throws Exception
     */
    public static ByteArrayOutputStream getReportFileByJsonArray(ExcelFile excelFile, JSONArray jsonArray) throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject jsonObject = jsonArray.getJSONObject(i);
                String name = jsonObject.getString("index");
                JSONArray users = jsonObject.getJSONArray("users");
                ExcelSheet excelSheet = excelFile.getExcelDoc().getExcelSheetMap().get("sheet");
                excelSheet.setSheetName(name);
                workbook = createkeySheet(workbook, users, excelSheet);
            }
            workbook.write(bos);
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception("文件导出出错");
        }
        return bos;
    }

    /**
     * 报表导出
     *
     * @param jsonObject
     * @return
     * @throws Exception `根据json字段生成多个sheet
     */
    public static ByteArrayOutputStream getReportFileByJsonObject(ExcelFile excelFile, JSONObject jsonObject) throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            Map<String, ExcelSheet> excelSheets = excelFile.getExcelDoc().getExcelSheetMap();
            Iterator<String> fields = jsonObject.keySet().iterator();
            while (fields.hasNext()) {
                String name = fields.next();
                JSONArray jsonArray = jsonObject.getJSONArray(name);
                ExcelSheet excelSheet = excelSheets.get(name);
                if (excelSheet != null) {
                    workbook = createDefaultSheet(workbook, jsonArray, excelSheet);
                }
            }
            workbook.write(bos);
        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception("文件导出出错");
        }
        return bos;
    }

    /**
     * 默认调用方法
     *
     * @param workbook
     * @param jsonArray
     * @param excelSheet
     * @return
     */
    private static XSSFWorkbook createkeySheet(XSSFWorkbook workbook, JSONArray jsonArray, ExcelSheet excelSheet) throws Exception {
        XSSFSheet sheet = workbook.createSheet(excelSheet.getSheetName());
        List<ExcelTitleTem> titleTems = excelSheet.getTitleTemList();
        sheet.setDisplayGridlines(false);
        XSSFCellStyle headStyle = ExcelFileUtil.getTitleCellStyle(workbook.createCellStyle(), workbook.createFont(), (short) 54);
        XSSFCellStyle intStyle = getContentCellStyle(workbook.createCellStyle(), workbook.createFont(), (short) 0);
        XSSFCellStyle stringStyle = (XSSFCellStyle) intStyle.clone();
        XSSFCellStyle doubleStyle = (XSSFCellStyle) intStyle.clone();
        intStyle.setDataFormat((short) 3);
        doubleStyle.setDataFormat((short) 4);
        XSSFRow tableHead = sheet.createRow(0);
        XSSFRow tableHead1 = sheet.createRow(1);
        for (int i = 0; i < titleTems.size(); i++) {
            XSSFCell cell = tableHead.createCell(i);
            cell.setCellValue(titleTems.get(i).getTitleName());
            cell.setCellStyle(headStyle);
            tableHead1.createCell(i).setCellStyle(headStyle);
            //合并左侧头部单元格
            sheet.addMergedRegion(new CellRangeAddress(0, 1, i, i));
        }
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject jsonData = jsonArray.getJSONObject(i);
            XSSFRow row = sheet.createRow(i + 2);
            for (int k = 0; k < titleTems.size(); k++) {
                String fileName = titleTems.get(k).getDateName();
                XSSFCell cell = row.createCell(k);
                setCellFormatValue(cell, jsonData.get(fileName), stringStyle, intStyle, doubleStyle);
            }
        }
        ExcelFileUtil.setAutoColumnSize(workbook, workbook.getSheetIndex(sheet), 0);
        return workbook;
    }

    /**
     * 默认调用方法
     *
     * @param workbook
     * @param jsonArray
     * @param excelSheet
     * @return
     */
    private static XSSFWorkbook createDefaultSheet(XSSFWorkbook workbook, JSONArray jsonArray, ExcelSheet excelSheet) throws Exception {
        XSSFSheet sheet = workbook.createSheet(excelSheet.getSheetName());
        List<ExcelTitleTem> titleTems = excelSheet.getTitleTemList();
        sheet.setDisplayGridlines(false);
        XSSFCellStyle headStyle = ExcelFileUtil.getTitleCellStyle(workbook.createCellStyle(), workbook.createFont(), (short) 54);
        XSSFCellStyle intStyle = getContentCellStyle(workbook.createCellStyle(), workbook.createFont(), (short) 0);
        XSSFCellStyle stringStyle = (XSSFCellStyle) intStyle.clone();
        XSSFCellStyle doubleStyle = (XSSFCellStyle) intStyle.clone();
        intStyle.setDataFormat((short) 3);
        doubleStyle.setDataFormat((short) 4);
        XSSFRow tableHead = sheet.createRow(0);
        XSSFRow tableHead1 = sheet.createRow(1);
        for (int i = 0; i < titleTems.size(); i++) {
            XSSFCell cell = tableHead.createCell(i);
            cell.setCellValue(titleTems.get(i).getTitleName());
            cell.setCellStyle(headStyle);
            tableHead1.createCell(i).setCellStyle(headStyle);
            //合并左侧头部单元格
            sheet.addMergedRegion(new CellRangeAddress(0, 1, i, i));
        }
        for (int i = 0; i < jsonArray.size(); i++) {
            JSONObject jsonData = jsonArray.getJSONObject(i);
            XSSFRow row = sheet.createRow(i + 2);
            for (int k = 0; k < titleTems.size(); k++) {
                String fileName = titleTems.get(k).getDateName();
                XSSFCell cell = row.createCell(k);
                setCellFormatValue(cell, jsonData.get(fileName), stringStyle, intStyle, doubleStyle);
            }
        }
        ExcelFileUtil.setAutoColumnSize(workbook, workbook.getSheetIndex(sheet), 0);
        return workbook;
    }

}
