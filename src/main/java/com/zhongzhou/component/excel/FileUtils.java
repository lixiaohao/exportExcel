package com.zhongzhou.component.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by lixiaohao on 2017/5/9
 *
 * @Description
 * @Create 2017-05-09 13:25
 * @Company
 */
public class FileUtils {

    public static byte[] getFileByte(File file){
        byte[] bytes = null;

      try {
          FileInputStream in = new FileInputStream(file);
          byte[] buff = new byte[1024];
          int i;
          ByteArrayOutputStream out = new ByteArrayOutputStream();
          while ((i = in.read(buff))>0) {
            out.write(buff,0,i);
          }

         bytes = out.toByteArray();
          in.close();
          out.close();
      }catch (IOException e){
          e.printStackTrace();
      }
        return bytes;
    }
}
