package com.zhongzhou.component.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
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
    /**标题*/
    private String title;
    /**第一个sheet名称*/
    private String sheetName;
    /**表头*/
    private Map<String,String> headRow;
    /** 需要导出的内容 */
    private List<Map<String,String>> values;
    /** 是否展示title */
    private Boolean showTitle = false;
    /**是否显示英文表头*/
    private Boolean showEnHead = false;
    /***
     * 展示标题行
     */
    public void showTitle () {
        this.showTitle = true;
    }
    /***
     * 展示英文表头
     */
    public void showEnHead(){
        this.showEnHead = true;
    }


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

    /***
     * 导出Excel到前端
     * @param response
     * @param fileName
     */
    public void write(HttpServletResponse response,String fileName) throws IOException {

        this.convertDataIntoExcel();

        String var = this.reMakeFileName(fileName);
        response.setHeader("Content-Disposition", "attachment;filename=" + var );
        try {
            OutputStream out = response.getOutputStream();
            wb.write( out );
            out.flush();
            out.close();
        }catch (IOException e) {
            e.printStackTrace();
            throw  new IOException("导出IO异常:",e.getCause());
        }

    }

    /***
     * 把Excel转化为byte[]
     * @return
     */
    public  byte[] importToByte(){
        convertDataIntoExcel();
       return toByte();
    }



    /**表头的总列数*/
    private int COLUMNS;
    /** 表头英文名称 */
    private String[]      HEAD;
    /** 后缀 */
    private final String    SUFFIX = ".xlsx";
    /** Excel默认名称 */
    private final String    DEFAULT_NAME = "import";
    /** 默认sheet名称 */
    private final String    DEFAULT_SHEET_NAMES[]= {"sheet1"};
    private XSSFWorkbook wb;

    private CellStyle titleStyle;        // 标题行样式
    private Font titleFont;              // 标题行字体
    private  CellStyle dateStyle;         // 日期行样式
    private  CellStyle headStyle;         // 表头行样式
    private  CellStyle contentStyle ;     // 内容行样式


    /***
     * 将数据转为Excel
     */
    private void convertDataIntoExcel() {
        this.validateParams();
        this.before();
        this.createHeadRow();
        this.fillContent();
    }


    private void validateParams() {
        if ( showTitle && null == title && "".equals(title)) {
            throw new  IllegalArgumentException("标题不能为空. title:"  + title);
        }
        if ( headRow == null) {
            throw new  IllegalArgumentException("表头不能为空. headRow:"  + headRow);
        }
        if (  headRow.size()==0 ) {
            throw new  IllegalArgumentException("表头内容不能为空. headRow.size():0");
        }

    }

    private void before(){

        Iterator<String> iterator = headRow.keySet().iterator();
        this.COLUMNS = headRow.size();

        String[] vars = new String[COLUMNS];
        int i = 0;
        while (iterator.hasNext()) {
            vars[i++] = iterator.next();
        }
        HEAD = vars;

        this.initWorkBook();
    }

    /***
     * 重构文件名
     * @param fileName
     * @return
     */
    private String reMakeFileName( String fileName ) {

        if ( fileName == null || "".equals(fileName) ) {
            fileName = DEFAULT_NAME+SUFFIX;
        }

        if ( !fileName.endsWith( SUFFIX ) ) {
            fileName += SUFFIX;
        }

        return fileName;
    }


    private void fillContent(){

        XSSFSheet sheetAt0 = wb.getSheetAt(0);
        if (values == null) {
            return;
        }

        int step = showEnHead?2:1;

        for (int v=0 ,len = values.size();v<len;v++) {

            Map<String,String> map = values.get(v);

            XSSFRow row  = sheetAt0.createRow( v + step );
            for (int i=0; i <COLUMNS;i++) {
                String value = map.get(HEAD[i]);
                XSSFCell cell = row.createCell(i);
                cell.setCellValue(value);
            }
        }


    }

    private void initWorkBook() {
        wb = new XSSFWorkbook();
        this.createSheet(sheetName);
    }

    private XSSFSheet createHeadRow(){

        int len = this.headRow.size();

        XSSFSheet sheet = wb.getSheetAt(0);

        this.createTitle(sheet,title,len);

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
            XSSFRow headRowEn = showTitle?sheet.createRow(2):sheet.createRow(1);
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
    private   byte[] toByte(){
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
     * @param sheetName
     */
    private void createSheet(String sheetName){
        String var = reMakeSheetName(sheetName);
        wb.createSheet( var );
    }

    private String reMakeSheetName( String sheetName){
        return  sheetName==null || "".equals(sheetName) ? DEFAULT_SHEET_NAMES[0]:sheetName;
    }



    /**
     * 创建 名称 ，需要合并单元格
     * @param sheet
     * @param title
     * @param length
     */
    private void createTitle(XSSFSheet sheet, String title, int length){

        if (!showTitle) {
            return;
        }

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


}
