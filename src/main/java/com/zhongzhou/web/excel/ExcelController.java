package com.zhongzhou.web.excel;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.zhongzhou.component.excel.ExcelPaser;
import com.zhongzhou.component.excel.ImportTemplate;
import com.zhongzhou.web.excel.vo.ExportVo;
import org.junit.Test;
import org.springframework.stereotype.Controller;
import org.springframework.validation.BindingResult;
import org.springframework.validation.ObjectError;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.context.ContextLoader;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.multipart.commons.CommonsMultipartResolver;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
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
        return "index.html";
    }

    @RequestMapping(value = "getModel",method = RequestMethod.POST)
    @ResponseBody
    public String getModel(
            @Valid @RequestBody ExportVo vo,
            BindingResult result,
            HttpServletRequest request,
            HttpServletResponse response){

        StringBuffer str = new StringBuffer("");
        if ( result.hasErrors() ) {
            List<ObjectError> errorList = result.getAllErrors();
            for(ObjectError error : errorList){
                str.append( error.getDefaultMessage() );
            }
            return str.toString();
        }

        String fileName = "template.xlsx";
        String path =  this.getClass().getClassLoader().getResource("").getPath();
        String realPath = path.substring(0,path.indexOf("excel")+6)+fileName;
        String sheetname = "firstsheet";

        ImportTemplate template = new ImportTemplate();
        template.createHeadRow(vo.getTitle(),sheetname,vo.getHeadRow());
        template.export(realPath);
        FileInputStream in = null;
        OutputStream out;
        File file = new File(realPath);
       try {

            in = new FileInputStream(file);
            out =  response.getOutputStream();

           response.reset();
           response.addHeader("Content-Disposition", "attachment;filename=" + new String(file.getName().getBytes()));
           byte[] buff = new byte[1024];
           int len = -1;
           while ( (len = in.read(buff))>0 ) {
               out.write(buff,0,len);
           }
           out.flush();
            in.close();
           out.close();

//           file.delete();
       }catch (FileNotFoundException e){
           e.printStackTrace();

       }catch (IOException e1){
           e1.printStackTrace();
       }finally {
       }

       return "success";
    }



    @RequestMapping(method=RequestMethod.POST, value="import")
    public  @ResponseBody JsonResponse upload(
            HttpServletRequest request) throws IOException{
        HttpSession session = request.getSession();

        JsonResponse 	response 		= new JsonResponse();
        boolean 		flag 			= true;
        StringBuilder 	actionMessage 	= new StringBuilder();

        MultipartFile file = null;

        List<Map<String,String>> datas = new ArrayList<Map<String, String>>();
        try {

            CommonsMultipartResolver commonsMultipartResolver = new
                    CommonsMultipartResolver(session.getServletContext());
            //设置编码
            commonsMultipartResolver.setDefaultEncoding("utf-8");
            //判断 request 是否有文件上传,即多部分请求...
            if (commonsMultipartResolver.isMultipart(request))
            {
                //转换成多部分request
                MultipartHttpServletRequest multipartRequest =
                        commonsMultipartResolver.resolveMultipart(request);

                // file 是指 文件上传标签的 name=值
                // 根据 name 获取上传的文件...
                file = multipartRequest.getFile("file");

            }

            if( file == null || file.isEmpty() ) {
                response.setActionMessage(" Cannot upload an empty file! ");
                return response;
            }

            String fileName = file.getOriginalFilename();
            String suffixName = fileName.substring(fileName.lastIndexOf(".") +1 );



            if ( !"xlsx".equals(suffixName)  && !"zip".equals(suffixName) ) {
                flag = false;
                actionMessage.append(" Only upload zip or xlsx type! please check it.your file name: " + fileName) ;
            }

            InputStream in = file.getInputStream();
            datas = ExcelPaser.readXlsx( in );

            response.setData(datas);
        } catch (Exception e) {
            // TODO: handle exception
            e.printStackTrace();
        }

        response.setActionMessage(actionMessage.toString());

        response.setSuccess(flag);

        return response;

    }

    @Test
    public void test(){
        ExportVo vo = new ExportVo();
        vo.setTitle("title");
        Map<String,String> headRow = new HashMap<String, String>();
        headRow.put("name","姓名");
        headRow.put("grade","年级");
        headRow.put("class","班级");
        headRow.put("score","分数");
        String title = "this is first sheet";
        vo.setHeadRow( headRow );

        Map<String,String> data1 = new HashMap<String, String>();
        data1.put("name","张三");
        data1.put("grade","3年级");
        data1.put("class","1班");
        data1.put("score","50");
        List<Map<String,String>> datas = new ArrayList<Map<String, String>>();
        datas.add(data1);
//        vo.setValues(datas);
        ObjectMapper mapper = new ObjectMapper();
        String json = null;
        try {
            json = mapper.writeValueAsString(vo);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
        System.out.println(json);
    }
}
