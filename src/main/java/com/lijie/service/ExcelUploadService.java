package com.lijie.service;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.lijie.excelFile.ExcelFile;
import com.lijie.excelFile.ExcelTitleTem;
import com.lijie.util.ExcelFileUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Objects;

@Service
public class ExcelUploadService {
    /**
     * 单个Sheet解析
     * @param workbook
     * @param file
     * @param dataRowNum
     * @return
     */
    public Object changeFile(XSSFWorkbook workbook,ExcelFile file,int dataRowNum){
        JSONArray jsonArray = new JSONArray();
        List<ExcelTitleTem> titleTems = file.getExcelDoc().getExcelSheetMap().get("sheet").getTitleTemList();
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rowNum = sheet.getPhysicalNumberOfRows();
        for (int i = dataRowNum; i < rowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            int colNum = row.getPhysicalNumberOfCells();
            JSONObject jsonObject = new JSONObject();
            for (int j = 0; j < colNum; j++) {
                XSSFCell cell = row.getCell(j);
                Object a = ExcelFileUtil.getCellValue(cell);
                jsonObject.put(titleTems.get(j).getImDateName(),a);
            }
            jsonArray.add(jsonObject);
        }
        return jsonArray;
    }



}
