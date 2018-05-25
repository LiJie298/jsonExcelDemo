package com.lijie.excelFile;

import java.util.List;

/**
 * @author lijie7
 * @date 2018/4/13
 * @Description
 * @modified By
 */
public class ExcelFileBase {
    /**
     * 文件类型 指定接口类型 获取 配置文件的 对应的 配置信息
     */
    protected int type;
    /**
     * 文件名
     */
    protected String name;
    protected List<ExcelSheet> sheets;

    public ExcelFileBase(int type, String name, List<ExcelSheet> sheets) {
        this.type = type;
        this.name = name;
        this.sheets = sheets;
    }
}
