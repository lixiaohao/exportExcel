package com.zhongzhou.web.excel.vo;

import java.util.List;
import java.util.Map;

/**
 * Created by lixiaohao on 2017/5/4
 *
 * @Description
 * @Create 2017-05-04 17:37
 * @Company
 */
public class ExportVo {
    /**标题*/
    private String title;
    /**表头 ，key:字段名，value:对应中文名*/
    private Map<String,String> headRow;

    /**每一行的值*/
    private List<Map<String,String>> values;

    public ExportVo() {
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Map<String, String> getHeadRow() {
        return headRow;
    }

    public void setHeadRow(Map<String, String> headRow) {
        this.headRow = headRow;
    }

    public List<Map<String, String>> getValues() {
        return values;
    }

    public void setValues(List<Map<String, String>> values) {
        this.values = values;
    }
}
