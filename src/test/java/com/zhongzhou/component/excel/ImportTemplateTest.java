package com.zhongzhou.component.excel;

import org.junit.Test;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by lixiaohao on 2017/5/4
 *
 * @Description
 * @Create 2017-05-04 16:55
 * @Company
 */
public class ImportTemplateTest {

    @Test
    public void exportTest(){

        Map<String,String> params = new HashMap<String, String>();
        params.put("name","姓名");
        params.put("grade","年级");
        params.put("class","班级");
        params.put("score","分数");

      /*  List<Map<String,String>> datas = new ArrayList<Map<String, String>>();
        Map<String,String> data1 = new HashMap<String, String>();
        params.put("name","张三");
        params.put("grade","3年级");
        params.put("class","1班");
        params.put("score","50");

        Map<String,String> data1 = new HashMap<String, String>();
        params.put("name","张三");
        params.put("grade","3年级");
        params.put("class","1班");
        params.put("score","50");*/


        String title = "this is first sheet";
        String sheetname = "firstsheet";

        ImportTemplate template = new ImportTemplate();
        template.createHeadRow(title,sheetname,params);
        template.export();
    }
}
