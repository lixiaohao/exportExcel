package com.zhongzhou.web.excel;

import com.zhongzhou.web.excel.vo.ExportVo;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;

/**
 * Created by lixiaohao on 2017/5/4
 *
 * @Description
 * @Create 2017-05-04 17:14
 * @Company
 */
@Controller
@RequestMapping("/excel/")
public class ExcelController {

    @RequestMapping("index")
    @ResponseBody
    public String index(){
        return "index.html";
    }

    @RequestMapping(value = "getModel",method = RequestMethod.POST)
    @ResponseBody
    public String getModel(
            @RequestBody ExportVo vo,
            HttpServletResponse response){

        String title = vo.getTitle();
        System.out.println("1234567");
        return "success";
    }
}
