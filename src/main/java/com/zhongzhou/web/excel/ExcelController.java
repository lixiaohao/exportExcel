package com.zhongzhou.web.excel;

import com.zhongzhou.component.excel.ImportTemplate;
import com.zhongzhou.web.excel.vo.ExportVo;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;
import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
        return "/WEB-INF/index.html";
    }



    @RequestMapping(value = "excel",method = RequestMethod.POST)
    public void export( HttpServletResponse response ){

        response.setHeader("Access-Control-Allow-Origin", "*");//解决跨域问题

        ExportVo vo = new ExportVo();
        vo.setTitle("测试一下");
        Map<String,String> headRow =  new HashMap<String, String>();
        headRow.put("name","姓名");
        headRow.put("age","年龄");
        headRow.put("height","身高");
        headRow.put("address","国籍");

        List<Map<String,String>> params = new ArrayList<Map<String, String>>();
        Map<String,String> dataMap =  new HashMap<String, String>();
        dataMap.put("name","刘德华");
        dataMap.put("age","15");
        dataMap.put("height","168厘米");
        dataMap.put("address","中国香港");
        vo.setHeadRow(headRow);
        params.add( dataMap );
        vo.setValues( params );

        ImportTemplate template = new ImportTemplate(headRow,params);
        template.showEnHead();
        template.setSheetName("sheetname");
        template.setTitle("标题");
        try {
            template.write(response,null);
        } catch (IOException e) {
            e.printStackTrace();
        }
//        byte[] fileByte  = template.export();
//
//        String fileName = "hehehhe.xlsx";
//
//        try {
//            response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
//            OutputStream out = response.getOutputStream();
//            out.write(fileByte,0,fileByte.length);
//
//            out.flush();
//            out.close();
//        }catch (IOException e) {
//            e.printStackTrace();
//        }

//        HttpHeaders header = new HttpHeaders();
//        String fileName = null;
//        ResponseEntity<byte[]> responseEntiry = null;
//        header.setContentType(MediaType.APPLICATION_OCTET_STREAM);
//        try{
//            fileName=new String("test.xlsx".getBytes("UTF-8"),"iso-8859-1");//为了解决中文名称乱码问题
//            header.setContentDispositionFormData("attachment", fileName);
//            responseEntiry = new ResponseEntity<byte[]>(fileByte,header, HttpStatus.CREATED);
//        }catch (UnsupportedEncodingException e){
//            e.printStackTrace();
//        }catch (IOException e1){
//            e1.printStackTrace();
//        }

    }


//
//
//    @RequestMapping(method=RequestMethod.POST, value="import")
//    public  @ResponseBody XmlResponse upload(
//            HttpServletRequest request) throws IOException{
//        HttpSession session = request.getSession();
//
//        XmlResponse 	response 		= new XmlResponse();
//        boolean 		flag 			= true;
//        StringBuilder 	actionMessage 	= new StringBuilder();
//
//        MultipartFile file = null;
//
//        List<Map<String,String>> datas = new ArrayList<Map<String, String>>();
//        try {
//
//            CommonsMultipartResolver commonsMultipartResolver = new
//                    CommonsMultipartResolver(session.getServletContext());
//            //设置编码
//            commonsMultipartResolver.setDefaultEncoding("utf-8");
//            //判断 request 是否有文件上传,即多部分请求...
//            if (commonsMultipartResolver.isMultipart(request))
//            {
//                //转换成多部分request
//                MultipartHttpServletRequest multipartRequest =
//                        commonsMultipartResolver.resolveMultipart(request);
//
//                // file 是指 文件上传标签的 name=值
//                // 根据 name 获取上传的文件...
//                file = multipartRequest.getFile("file");
//
//            }
//
//            if( file == null || file.isEmpty() ) {
//                response.setActionMessage(" Cannot upload an empty file! ");
//                return response;
//            }
//
//            String fileName = file.getOriginalFilename();
//            String suffixName = fileName.substring(fileName.lastIndexOf(".") +1 );
//
//
//
//            if ( !"xlsx".equals(suffixName)  && !"zip".equals(suffixName) ) {
//                flag = false;
//                actionMessage.append(" Only upload zip or xlsx type! please check it.your file name: " + fileName) ;
//            }
//
//            InputStream in = file.getInputStream();
//            datas = ExcelPaser.readXlsx( in );
//
//            response.setData(datas);
//        } catch (Exception e) {
//            // TODO: handle exception
//            e.printStackTrace();
//        }
//
//        response.setActionMessage(actionMessage.toString());
//
//        response.setSuccess(flag);
//
//        return response;
//
//    }

}
