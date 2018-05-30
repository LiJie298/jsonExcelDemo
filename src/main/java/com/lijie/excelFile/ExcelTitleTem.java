package com.lijie.excelFile;

/**
 * @author lijie7
 * @date 2018/4/13
 * @Description
 * @modified By
 * 导出或者导入模板的表头信息
 */
public class ExcelTitleTem {
    /**
     * 文件标题字段
     */
    private String titleName;
    /**
     * json -> excel 时对应的json的key值
     */
    private String dateName;
    /**
     * excel -> json 时 对应的json字段
     */
    private String imDateName;

    public ExcelTitleTem() {
    }

    public ExcelTitleTem(String titleName, String dateName, String imDateName) {
        this.titleName = titleName;
        this.dateName = dateName;
        this.imDateName = imDateName;
    }

    public String getTitleName() {
        return titleName;
    }

    public void setTitleName(String titleName) {
        this.titleName = titleName;
    }

    public String getDateName() {
        return dateName;
    }

    public void setDateName(String dateName) {
        this.dateName = dateName;
    }

    public String getImDateName() {
        return imDateName;
    }

    public void setImDateName(String imDateName) {
        this.imDateName = imDateName;
    }
}
