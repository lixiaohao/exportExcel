package com.zhongzhou.component.pdf;

import com.itextpdf.io.font.FontConstants;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.*;
import org.junit.Test;

import javax.print.DocFlavor;
import java.io.*;
import java.util.*;
import java.util.List;

import static com.itextpdf.kernel.pdf.PdfName.BaseFont;


/**
 * Created by lixiaohao on 2017/5/10
 *
 * @Description
 * @Create 2017-05-10 9:52
 * @Company
 */
public class PDFUtils {

	private Map<String,String> headRow;
	private String[] head;
	private int      columnNum;

	private Document document;
	private Table table;

	private int width =	3;
	private float[] columnWidths;

	PdfFont font = PdfFontFactory.createFont(FontConstants.HELVETICA);
//	BaseFont baseFont;
	String dest = "C:\\Users\\lixiaohao.ZZGRP\\Desktop\\temp\\pdf\\test.pdf";

	public PDFUtils(Map<String, String> headRow) throws IOException {
		this.headRow = headRow;
		prePdf();
	}

	public byte[] creataPdf(List<Map<String,String>> data){
		byte[] pdfByte = null;

		createTable();
		createHeader();
		filDataTobody(data);

		document.add(table);
		document.close();
		return pdfByte;
	}

	private void prePdf(){
		if ( headRow ==  null ) {
			throw new IllegalArgumentException(" headRow is null,please check it! ");
		}
		columnNum = headRow.size();
		head = new String[columnNum];
		Set<String> set = headRow.keySet();
		Iterator<String> iterator = set.iterator();
		int i = 0;
		while (iterator.hasNext()) {
			head[i++] = iterator.next();
		}
//			PdfWriter writer = new PdfWriter(new ByteArrayOutputStream());
		try {

			PdfWriter writer = new PdfWriter(dest);
			PdfDocument pdf = new PdfDocument(writer);
			document = new Document(pdf, PageSize.A4.rotate());
			document.setMargins(20, 20, 20, 20);
		}catch (FileNotFoundException e){
			e.printStackTrace();
		}catch (IOException e1){
			e1.printStackTrace();
		}

	}


	private void createTable(){
		if ( columnWidths == null ) {
			columnWidths =  new float[columnNum];
		}

		for (int i=0; i<columnNum ;i++) {
			columnWidths[i] = width;
		}

		table = new Table(columnWidths);
		table.setWidthPercent(100);
	}

	/***
	 * 创建头
	 */
	private void createHeader(){
		//first row
		process(table,head,true);
		String[] values = getMapValueToArray(headRow);
		//second row
		process(table,values,false);
	}

	/***
	 * 将数据填充到PDF中
	 * @param data
	 */
	private void filDataTobody(java.util.List<Map<String,String>> data){
		for (Map<String,String> map:data) {
			String[] values = getMapValueToArray(map);
			process(table,values,false);
		}
	}

	private String[] getMapValueToArray(Map<String,String> map){
		int len = columnNum;
		String[] values = new String[len];

		for (int i=0; i<len;i++) {
			values[i] = map.get(head[i]);
		}
		return values;
	}

	private void process(Table table, String[] line, boolean isHeader) {
		for (String cellValue:line) {
			if (isHeader) {
				table.addHeaderCell(
						new Cell().add(
								new Paragraph(cellValue).setFont(font)));
			} else {
				table.addCell(
						new Cell().add(
								new Paragraph(cellValue).setFont(font)));

			}
		}

		/*StringTokenizer tokenizer = new StringTokenizer(line, ";");
		while (tokenizer.hasMoreTokens()) {
			if (isHeader) {
				table.addHeaderCell(
						new Cell().add(
								new Paragraph(tokenizer.nextToken()).setFont(font)));
			} else {
				table.addCell(
						new Cell().add(
								new Paragraph(tokenizer.nextToken()).setFont(font)));
			}
		}*/
	}
}
