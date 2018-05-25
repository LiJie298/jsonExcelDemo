package com.lijie.excelFile;

import java.util.List;

/**
 * @author lijie7
 * @date 2018/5/22
 * @Description
 * @modified By
 */
public class ExcelSheet {
    /**
     * 工作表名称
     */
    private  String sheetName;

    /**
     * 工作表中表头跟 json key 相对的数据
     */
    private List<ExcelTitleTem> titleTemList;

    public ExcelSheet(String sheetName, List<ExcelTitleTem> titleTemList) {
        this.sheetName = sheetName;
        this.titleTemList = titleTemList;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<ExcelTitleTem> getTitleTemList() {
        return titleTemList;
    }

    public void setTitleTemList(List<ExcelTitleTem> titleTemList) {
        this.titleTemList = titleTemList;
    }
}
