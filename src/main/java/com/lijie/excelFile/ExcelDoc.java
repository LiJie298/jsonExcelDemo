package com.lijie.excelFile;

import java.util.Map;

/**
 * @author lijie7
 * @date 2018/4/13
 * @Description
 * @modified By
 */
public class ExcelDoc {

    /**
     * 工作表名称
     * 工作表中表头跟 json key 相对的数据
     */
    private Map<String, ExcelSheet> excelSheetMap;


    public ExcelDoc() {
    }

    public ExcelDoc(Map<String, ExcelSheet> excelSheetMap) {
        this.excelSheetMap = excelSheetMap;
    }

    public Map<String, ExcelSheet> getExcelSheetMap() {
        return excelSheetMap;
    }

    public void setExcelSheetMap(Map<String, ExcelSheet> excelSheetMap) {
        this.excelSheetMap = excelSheetMap;
    }
}
