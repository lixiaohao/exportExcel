package com.zhongzhou.component.excel;

import com.zhongzhou.web.excel.vo.ExportVo;
import org.junit.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
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

        String path = "C:\\Users\\lixiaohao.ZZGRP\\Desktop\\temp\\excel\\test.xlsx";
        ExportVo vo = getVo();
        ImportTemplate template = new ImportTemplate(vo);
        template.createHeadRow(vo.getTitle(),"first Sheet",vo.getHeadRow());
        template.fillContent(vo.getValues());
        template.exportExcel();
    }


    public ExportVo getVo(){
        ExportVo vo = new ExportVo();
        vo.setTitle("title");
        Map<String,String> headRow = new HashMap<String, String>();
        headRow.put("name","姓名");
        headRow.put("grade","年级");
        headRow.put("class","班级");
        headRow.put("score","分数");
        vo.setHeadRow( headRow );

        List<Map<String,String>> datas = new ArrayList<Map<String, String>>();
        Map<String,String> data1= new HashMap<String, String>();
        data1.put("name","刘德华");
        data1.put("grade","三年级");
        data1.put("class","2班");
        data1.put("score","50");
        datas.add(data1);

        Map<String,String> data2= new HashMap<String, String>();
        data2.put("name","周星驰");
        data2.put("grade","两年级");
        data2.put("class","1班");
        data2.put("score","70");
        datas.add(data1);
        vo.setValues( datas );

        return vo;
    }

}
