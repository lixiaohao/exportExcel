package com.zhongzhou.component.excel;

import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by lixiaohao on 2017/5/5
 *
 * @Description
 * @Create 2017-05-05 16:34
 * @Company
 */
public class ExcelPaser {
    /**
     * Read the Excel 2010+
     * 仅支持字符串类型
     * @param
     * @return
     * @throws IOException
     */
    public static List<Map<String,String>> readXlsx(InputStream is) {

        List<Map<String,String>> excelValues = new ArrayList<Map<String, String>>();

        String exceptionField = "";
        Integer    exceptionRowNum= null;

        try {
            Map<String, String> excelValue ;

            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);

            //仅支持一个sheet
            if(xssfWorkbook.getNumberOfSheets() <= 0 )
                return excelValues;

            XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);

            if (xssfSheet == null)
                return excelValues;

            // Read the Row
            int lastRowNum = xssfSheet.getLastRowNum();

            if(lastRowNum<=1)
                return excelValues;

            XSSFRow headRow = xssfSheet.getRow(2);

            String[] headArgs = new String[headRow.getLastCellNum()+1];

            for(int cellNum = 0 ;cellNum<headRow.getLastCellNum(); cellNum++){
                XSSFCell cell = headRow.getCell(cellNum);
                String value = "";
                if(cell != null){
                    value = (cell.getStringCellValue()==null)?"":cell.getStringCellValue().trim();
                }
                headArgs[cellNum] = value;
            }


            for (int rowNum = 3; rowNum <= lastRowNum; rowNum++) {

                XSSFRow xssfRow = xssfSheet.getRow(rowNum);

                if (xssfRow != null) {

                    excelValue = new HashMap<String, String>();
                    for(int cellIndex = 0;cellIndex < headArgs.length-1;cellIndex++){


                        exceptionField = headArgs[cellIndex];
                        exceptionRowNum= rowNum;

                        XSSFCell nos = xssfRow.getCell(cellIndex);

                        if(nos  == null)
                            continue;

                        String value = getCellValue(nos);
                        if( value == null || value.trim().equals("") ){
                            continue;
                        }
                        excelValue.put(headArgs[cellIndex],value);
                    }

                    if( excelValue.size()<=0 ){
                        continue;
                    }

                    excelValues.add(excelValue);

                }
            }
        }catch (Exception e){
            e.printStackTrace();
            System.out.println("错误行字段:"+ exceptionField+"  行号:"+exceptionRowNum+3);
        }

        return excelValues;

    }

    /**
     * 仅支持  文本类型 和 数值类型（如果为数值类型，则转化为字符串）
     * @param xssfCell
     * @return
     */
    private static String getCellValue(XSSFCell xssfCell){

        String value = null;

        if( xssfCell == null )
        {
            value = null;
            return value;
        }

        if( xssfCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC ){

            HSSFDataFormatter dataFormatter = new HSSFDataFormatter();
            value = dataFormatter.formatCellValue(xssfCell);
        }else {
            value = (xssfCell.getStringCellValue() == null)?null:xssfCell.getStringCellValue().trim();
        }

//          switch ( xssfCell.getCellType() ){
//              case XSSFCell.CELL_TYPE_STRING :
//                  value = (xssfCell.getStringCellValue() == null)?null:xssfCell.getStringCellValue().trim();
//                  break;
//              default:
//                  throw new IllegalStateException("Cannot get a STRING value from a "+xssfCell.getCellTypeEnum()+" cell");
//          }

        return value;
    }


}
