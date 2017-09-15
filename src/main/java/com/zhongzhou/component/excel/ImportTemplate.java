package com.zhongzhou.component.excel;

import com.zhongzhou.web.excel.vo.ExportVo;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * Created by lixiaohao on 2017/5/4
 *
 * @Description
 * @Create 2017-05-04 16:53
 * @Company
 */
public class ImportTemplate {
    /**表头*/
    private String sheetName;   //第一个sheet名称
    private Map<String,String> headRow;

    private String title;           //标题
    private List<Map<String,String>> values;
    private int                columnNum; //表头的列数
    private String[]            head;     //表头英文名称
    private Boolean showTitle = true;      //是否有title
    private Boolean showEnHead = false;   //默认不显示英文名
    private ExportVo            vo;

    private XSSFWorkbook wb;

    private CellStyle titleStyle;        // 标题行样式
    private Font titleFont;              // 标题行字体
    private  CellStyle dateStyle;         // 日期行样式
    private  CellStyle headStyle;         // 表头行样式
    private  CellStyle contentStyle ;     // 内容行样式


    public ImportTemplate(String sheetName, Map<String, String> headRow, String title, List<Map<String, String>> values, Boolean showTitle, Boolean showEnHead) {
        this.sheetName = sheetName;
        this.headRow = headRow;
        this.title = title;
        this.values = values;
        this.showTitle = showTitle;
        this.showEnHead = showEnHead;
    }

    public ImportTemplate(Map<String, String> headRow, List<Map<String, String>> values) {
        this.headRow = headRow;
        this.values = values;
    }


    private void before(){
        wb = new XSSFWorkbook();

        Iterator<String> iterator = headRow.keySet().iterator();
        this.columnNum = headRow.size();

        String[] vars = new String[columnNum];
        int i = 0;
        while (iterator.hasNext()) {
            vars[i++] = iterator.next();
        }
        head = vars;
    }


//
//    public  byte[] creatExcel(){
//
//        try {
//            ByteArrayOutputStream out = new ByteArrayOutputStream();
//            wb.write(out);
//            wb.close();
//            return out.toByteArray();
//        }catch (Exception e){
//            e.printStackTrace();
//        }
//        return null;
//    }
//

    public  byte[] export(){

        if ( showTitle && null == title && "".equals(title)) {
            throw new  IllegalArgumentException("标题不能为空. title:"  + title);
        }
        before();
        createHeadRow();
        fillContent();
       return exportExcel();
    }

    private void fillContent(){
        if (values == null) {
            throw new IllegalArgumentException(" Invalid param\" values \"： "+values+",place check it.");
        }

        int step = showEnHead?2:1;

        for (int v=0 ,len = values.size();v<len;v++) {

            Map<String,String> map = values.get(v);

            XSSFRow row  = wb.getSheetAt(0).createRow( v + step );
            for (int i=0; i <columnNum;i++) {
                String value = map.get(head[i]);
                XSSFCell cell = row.createCell(i);
                cell.setCellValue(value);
            }
        }


    }
    private XSSFSheet createHeadRow(){
        this.headRow = headRow;

        int len = this.headRow.size();

        createSheet(sheetName);

        XSSFSheet sheet = wb.getSheetAt(0);

        if (null != title && !"".equals(title)) {
            createTitle(sheet,title,len);
            showTitle = true;
        }else {
            showTitle = false;
        }

        XSSFRow headRowZh = showTitle?sheet.createRow(1):sheet.createRow(0);
        int i = 0;
        for ( Map.Entry<String,String> entry:this.headRow.entrySet() ) {

            String keyZh = entry.getValue();
            XSSFCell cellZh = headRowZh.createCell( i );
            cellZh.setCellStyle( headStyle );
            cellZh.setCellValue( keyZh );

            if ( ++i == len ) {
                break;
            }
        }

        if ( showEnHead ) {
            XSSFRow headRowEn = showEnHead?sheet.createRow(2):sheet.createRow(1);
            int j = 0;
            for ( Map.Entry<String,String> entry:this.headRow.entrySet() ) {
                String keyEn = entry.getKey();
                XSSFCell cellEn = headRowEn.createCell( j );
                cellEn.setCellStyle( headStyle );
                cellEn.setCellValue( keyEn );

                if ( ++j == len ) {
                    break;
                }
            }
        }

        return sheet;
    }
    public  byte[] exportExcel(){
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);
            wb.close();
            return out.toByteArray();
        }catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }


    /**
     * 创建一个sheet，默认名称为 ：Sheet1
     * @param
     * @param sheetname
     */
    private void createSheet(String sheetname){
        wb.createSheet( (sheetname==null || sheetname.equals(""))?"Sheet1":sheetname );
    }

    /**
     * 创建 名称 ，需要合并单元格
     * @param sheet
     * @param title
     * @param length
     */
    private void createTitle(XSSFSheet sheet, String title, int length){
        CellRangeAddress titleRange = new CellRangeAddress(0, 0, 0, length-1);
        sheet.addMergedRegion(titleRange);
        XSSFRow titleRow = sheet.createRow(0);
        titleRow.setHeight((short) 800);
        XSSFCell titleCell = titleRow.createCell(0);
        titleCell.setCellStyle(titleStyle);
        titleCell.setCellValue( title );
    }

//    public XSSFSheet createHeadRow(String title,String sheetname,Map<String,String> headRow){
//        this.headRow = headRow;
//
//        int len = this.headRow.size();
//
//        createSheet(sheetname);
//        XSSFSheet sheet = wb.getSheetAt(0);
//
//        if (null != title && !"".equals(title)) {
//            createTitle(sheet,title,len);
//            showTitle = true;
//        }else {
//            showTitle = false;
//        }
//
//        XSSFRow headRowZh = showTitle?sheet.createRow(1):sheet.createRow(0);
//        int i = 0;
//        for ( Map.Entry<String,String> entry:this.headRow.entrySet() ) {
//
//
//            String keyZh = entry.getValue();
//            XSSFCell cellZh = headRowZh.createCell( i );
//            cellZh.setCellStyle( headStyle );
//            cellZh.setCellValue( keyZh );
//
//
////            XSSFCell cellEn = headRowEn.createCell( i );
////            cellEn.setCellStyle( headStyle );
////            cellEn.setCellValue( keyEn );
//
//            if ( ++i == len ) {
//                break;
//            }
//        }
//
//        if ( showEnHead ) {
//            XSSFRow headRowEn = showEnHead?sheet.createRow(2):sheet.createRow(1);
//            int j = 0;
//            for ( Map.Entry<String,String> entry:this.headRow.entrySet() ) {
//                String keyEn = entry.getKey();
//                XSSFCell cellEn = headRowEn.createCell( j );
//                cellEn.setCellStyle( headStyle );
//                cellEn.setCellValue( keyEn );
//
//                if ( ++j == len ) {
//                    break;
//                }
//            }
//        }
//
//        return sheet;
//    }


//    public void fillContent(List<Map<String,String>> values){
//        if (values == null) {
//            throw new IllegalArgumentException(" Invalid param\" values \"： "+values+",place check it.");
//        }
//
//        int step = showEnHead?2:1;
//
//        for (int v=0 ,len = values.size();v<len;v++) {
//
//            Map<String,String> map = values.get(v);
//
//            XSSFRow row  = wb.getSheetAt(0).createRow( v + step );
//            for (int i=0; i <columnNum;i++) {
//                String value = map.get(head[i]);
//                XSSFCell cell = row.createCell(i);
//                cell.setCellValue(value);
//            }
//        }
//
//
//    }

    /**
     * @Description: 初始化标题行样式
     */
    private void initTitleCellStyle()
    {
        titleStyle = wb.createCellStyle();
        //文字居中
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
//        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        titleStyle.setFont(titleFont);
        titleStyle.setFillBackgroundColor(IndexedColors.SKY_BLUE.index);
    }
    /**
     * @Description: 初始化日期行样式
     */
    private void initDateCellStyle()
    {
        dateStyle.setAlignment(CellStyle.ALIGN_CENTER_SELECTION);
        dateStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        dateStyle.setFont(dateFont);
        dateStyle.setFillBackgroundColor(IndexedColors.SKY_BLUE.index);
    }

    /**
     * @Description: 初始化表头行样式
     */
    private void initHeadCellStyle()
    {
        headStyle = wb.createCellStyle();
        headStyle.setAlignment(CellStyle.ALIGN_CENTER);
        headStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        headStyle.setFont(headFont);
        headStyle.setFillBackgroundColor(IndexedColors.YELLOW.index);
        headStyle.setBorderTop(CellStyle.BORDER_MEDIUM);
        headStyle.setBorderBottom(CellStyle.BORDER_THIN);
        headStyle.setBorderLeft(CellStyle.BORDER_THIN);
        headStyle.setBorderRight(CellStyle.BORDER_THIN);
        headStyle.setTopBorderColor(IndexedColors.BLUE.index);
        headStyle.setBottomBorderColor(IndexedColors.BLUE.index);
        headStyle.setLeftBorderColor(IndexedColors.BLUE.index);
        headStyle.setRightBorderColor(IndexedColors.BLUE.index);
    }

    /**
     * @Description: 初始化内容行样式
     */
    private void initContentCellStyle()
    {
        contentStyle.setAlignment(CellStyle.ALIGN_CENTER);
        contentStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
//        contentStyle.setFont(contentFont);
        contentStyle.setBorderTop(CellStyle.BORDER_THIN);
        contentStyle.setBorderBottom(CellStyle.BORDER_THIN);
        contentStyle.setBorderLeft(CellStyle.BORDER_THIN);
        contentStyle.setBorderRight(CellStyle.BORDER_THIN);
        contentStyle.setTopBorderColor(IndexedColors.BLUE.index);
        contentStyle.setBottomBorderColor(IndexedColors.BLUE.index);
        contentStyle.setLeftBorderColor(IndexedColors.BLUE.index);
        contentStyle.setRightBorderColor(IndexedColors.BLUE.index);
        contentStyle.setWrapText(true); // 字段换行
    }

    /**
     * @Description: 初始化标题行字体
     */
    private void initTitleFont()
    {
        titleFont.setFontName("华文楷体");
        titleFont.setFontHeightInPoints((short) 20);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        titleFont.setCharSet(Font.DEFAULT_CHARSET);
        titleFont.setColor(IndexedColors.BLUE_GREY.index);
    }

    /**
     * @Description: 初始化日期行字体
     */
    private void initDateFont()
    {
//        Font dateFont =
//        dateFont.setFontName("隶书");
//        dateFont.setFontHeightInPoints((short) 10);
//        dateFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
//        dateFont.setCharSet(Font.DEFAULT_CHARSET);
//        dateFont.setColor(IndexedColors.BLUE_GREY.index);
    }

    /**
     * @Description: 初始化表头行字体
     */
    private void initHeadFont()
    {
//        headFont.setFontName("宋体");
//        headFont.setFontHeightInPoints((short) 10);
//        headFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
//        headFont.setCharSet(Font.DEFAULT_CHARSET);
//        headFont.setColor(IndexedColors.BLUE_GREY.index);
    }

    /**
     * @Description: 初始化内容行字体
     */
    private void initContentFont()
    {
//        contentFont.setFontName("宋体");
//        contentFont.setFontHeightInPoints((short) 10);
//        contentFont.setBoldweight(Font.BOLDWEIGHT_NORMAL);
//        contentFont.setCharSet(Font.DEFAULT_CHARSET);
//        contentFont.setColor(IndexedColors.BLUE_GREY.index);
    }


    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Map<String, String> getHeadRow() {
        return headRow;
    }

    public void setHeadRow(Map<String, String> headRow) {
        this.headRow = headRow;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<Map<String, String>> getValues() {
        return values;
    }

    public void setValues(List<Map<String, String>> values) {
        this.values = values;
    }

    public Boolean getShowTitle() {
        return showTitle;
    }

    public void setShowTitle(Boolean showTitle) {
        this.showTitle = showTitle;
    }

    public Boolean getShowEnHead() {
        return showEnHead;
    }

    public void setShowEnHead(Boolean showEnHead) {
        this.showEnHead = showEnHead;
    }


}
