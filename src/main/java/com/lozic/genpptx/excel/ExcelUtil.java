package com.lozic.genpptx.excel;


import jxl.Workbook;
import jxl.format.UnderlineStyle;
import jxl.write.*;

import java.io.File;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Properties;

public class ExcelUtil {

	public static final int HEADERL = 0, //Excel头样式左
		HEADERC = 1, //Excel头样式中
		HEADERR = 2, //Excel头样式右
		TITLE = 3, //Excel表标题样式
		CONTENT = 4, //Excel表内容样式 
		HIDDEN = 5, //隐藏行和列的样式
		TABLE = 6; //表头样式

	WritableWorkbook wwb = null; //Excel工作薄的对象
	WritableSheet ws = null; //Sheet表对象
	//Cell cell = null ;//单元格对象

	// 到出到流中
	public ExcelUtil(OutputStream os) {
		try {
			wwb = Workbook.createWorkbook(os);
		} catch (Exception e) {
			System.out.println("ExcelUtil(String filesource) error=" + e);
		}

	}

	// 导出到文件中
	public ExcelUtil(String filesource) {
		try {
			wwb = Workbook.createWorkbook(new File(filesource));
		} catch (Exception e) {
			System.out.println("ExcelUtil(String filesource) error=" + e);
		}
    }

	// 导出到文件中
	public ExcelUtil(File fs) {
		try {
			wwb = Workbook.createWorkbook(fs);
		} catch (Exception e) {
			System.out.println("ExcelUtil(File fs) error=" + e);
		}
	}

	public void exp(ArrayList fieldNames, ArrayList data) {
		exp("", fieldNames, data);
	}
	
	public void makeExcel(String title, ArrayList fieldNames, ArrayList data, ArrayList l){
        	exp(title,fieldNames,data);
	}

	
	
	public void exp(String title, ArrayList fieldNames, ArrayList data) {
		ws = wwb.createSheet("MySheet1", 0);
		Properties prop = null;

		int iRow = (title != null && !"".equals(title)) ? 1 : 0; // 第几行

		try {
			//ws.setRowView(0,600);
			if (iRow == 1) {
				
				jxl.write.Label labelC = new jxl.write.Label(0, 0, title,getWritableCellFormat(ExcelUtil.TITLE));
				//ws.setRowView(1, 600);
				ws.addCell(labelC);
				ws.mergeCells(0, 0, fieldNames.size() - 1, 0);
			}

			for (int i = 0; i < data.size(); i++) {
				prop = (Properties) data.get(i);
				for (int j = 0; j < fieldNames.size(); j++) {
					String fieldName = fieldNames.get(j).toString();
					Object o = prop.get(fieldName);
					String sData = (o != null) ? o.toString() : "";
					jxl.write.Label labelC = null;
					if (i==0){
						labelC = new jxl.write.Label(j, i + iRow, sData,getWritableCellFormat(ExcelUtil.TABLE));
					}else{
						labelC = new jxl.write.Label(j, i + iRow, sData,getWritableCellFormat(ExcelUtil.CONTENT));
						
					}					
					ws.addCell(labelC);
					
				}					
					//ws.setRowView(i+1, 500);			
			}
			
			
			for (int k = 0; k < fieldNames.size(); k++) {
				ws.setColumnView(k, 24);
			}			
			 wwb.write();
			 wwb.close();

		} catch (Exception e) {
			System.out.println("exp error=" + e);
		}
	}

	public void addFiledsAndDesc(
		ArrayList fieldNames,
		String filedName,
		ArrayList data,
		String fieldDesc) {
		fieldNames.add(filedName);
		int position = data.size() - 1;
		if (position >= 0) {
			Properties prop = (Properties) data.get(position);
			prop.put(filedName, fieldDesc);
			data.set(position, prop);
		} else {
			Properties prop = new Properties();
			prop.put(filedName, fieldDesc);
			data.add(0, prop);
		}
	}

	public WritableCellFormat getWritableCellFormat(int type)	throws WriteException {
		WritableCellFormat detFormat = null;
		if (type == ExcelUtil.HEADERL) {
			WritableFont detFont =
				new WritableFont(
					WritableFont.ARIAL,
					10,
					WritableFont.NO_BOLD,
					false,
					UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			detFormat = new WritableCellFormat(detFont);
			detFormat.setWrap(true);
			detFormat.setAlignment(jxl.format.Alignment.LEFT);
			detFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			return detFormat;
		} else if (type == ExcelUtil.HEADERC) {
			WritableFont detFont =
				new WritableFont(
					WritableFont.ARIAL,
					10,
					WritableFont.NO_BOLD,
					false,
					UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			detFormat = new WritableCellFormat(detFont);
			detFormat.setWrap(true);
			detFormat.setAlignment(jxl.format.Alignment.CENTRE);
			detFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			return detFormat;
		} else if (type == ExcelUtil.HEADERR) {
			WritableFont detFont =
				new WritableFont(
					WritableFont.ARIAL,
					10,
					WritableFont.NO_BOLD,
					false,
					UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			detFormat = new WritableCellFormat(detFont);
			detFormat.setWrap(true);
			detFormat.setAlignment(jxl.format.Alignment.RIGHT);
			detFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			return detFormat;
		} else if (type == ExcelUtil.TITLE) {
			WritableFont detFont =
				new WritableFont(
					WritableFont.ARIAL,
					12,
					WritableFont.BOLD,
					false,
					UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.RED);
			detFormat = new WritableCellFormat(detFont);
			detFormat.setAlignment(jxl.format.Alignment.CENTRE);
			detFormat.setVerticalAlignment(
				jxl.format.VerticalAlignment.CENTRE);
			return detFormat;
		} else if (type == ExcelUtil.HIDDEN) {
			WritableFont detFont =
				new WritableFont(
					WritableFont.ARIAL,
					10,
					WritableFont.NO_BOLD,
					false,
					UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			detFormat = new WritableCellFormat(detFont);
			detFormat.setWrap(true);
			detFormat.setBorder(
				jxl.format.Border.ALL,
				jxl.format.BorderLineStyle.THIN);
			detFormat.setAlignment(jxl.format.Alignment.CENTRE);
			detFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			return detFormat;
		} else if (type == ExcelUtil.CONTENT) {
			WritableFont detFont =
				new WritableFont(
					WritableFont.ARIAL,
					10,
					WritableFont.NO_BOLD,
					false,
					UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			detFormat = new WritableCellFormat(detFont);
			detFormat.setWrap(true);
			detFormat.setBorder(
				jxl.format.Border.ALL,
				jxl.format.BorderLineStyle.THIN);
			detFormat.setAlignment(jxl.format.Alignment.CENTRE);
			detFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			//detFormat.setBackground(jxl.format.Colour.GREY_25_PERCENT);
			return detFormat;
		} else if (type == ExcelUtil.TABLE) {
			WritableFont detFont =
				new WritableFont(
					WritableFont.ARIAL,
					10,
					WritableFont.BOLD,
					false,
					UnderlineStyle.NO_UNDERLINE,
					jxl.format.Colour.BLACK);
			detFormat = new WritableCellFormat(detFont);
			detFormat.setWrap(true);
			detFormat.setBorder(
				jxl.format.Border.ALL,
				jxl.format.BorderLineStyle.THIN);
			detFormat.setAlignment(jxl.format.Alignment.CENTRE);
			detFormat.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);
			detFormat.setBackground(jxl.format.Colour.GREY_25_PERCENT);
			return detFormat;
		}
		
		return detFormat;
	}
	
	
	

	public WritableSheet getWs() {
		return ws;
	}

	public void setWs(WritableSheet ws) {
		this.ws = ws;
	}

	public WritableWorkbook getWwb() {
		return wwb;
	}

	public void setWwb(WritableWorkbook wwb) {
		this.wwb = wwb;
	}

	public static void main(String[] args) {
		

   	    ExcelUtil ew = new ExcelUtil("e:\\test.xls");
   		ArrayList fieldNames = new ArrayList();
		ArrayList data = new ArrayList();
		//生成表头
																	

		ew.addFiledsAndDesc(fieldNames, "Serial", data, "序号");
		ew.addFiledsAndDesc(fieldNames, "itemCode", data, "采购品编码");
		ew.addFiledsAndDesc(fieldNames, "itemName", data, "采购品名称");
		ew.addFiledsAndDesc(fieldNames, "spec", data, "规格");
		ew.addFiledsAndDesc(fieldNames, "unit", data, "单位");
		ew.addFiledsAndDesc(fieldNames, "num", data, "数量");
		ew.addFiledsAndDesc(fieldNames, "maker", data, "制造商");
		
		ew.addFiledsAndDesc(fieldNames, "pcode", data, "货号");
		ew.addFiledsAndDesc(fieldNames, "planprice", data, "计划价(元)");
		ew.addFiledsAndDesc(fieldNames, "applydept", data, "使用单位");
		ew.addFiledsAndDesc(fieldNames, "catalogcode", data, "所属分类编码");
		ew.addFiledsAndDesc(fieldNames, "applydate", data, "需求日期");
		ew.addFiledsAndDesc(fieldNames, "sellerName", data, "供应商名称");
		ew.addFiledsAndDesc(fieldNames, "price", data, "单价");
		
		ew.addFiledsAndDesc(fieldNames, "bz", data, "币种");
		ew.addFiledsAndDesc(fieldNames, "hsprice", data, "换算价（元）");
		ew.addFiledsAndDesc(fieldNames, "status", data, "状态");
		
		ew.exp("表头 ", fieldNames, data); 
    	 
    
		 
	}
 
}



