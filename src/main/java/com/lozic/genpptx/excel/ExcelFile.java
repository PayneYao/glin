package com.lozic.genpptx.excel;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class ExcelFile {

	public ExcelFile() {
		// TODO Auto-generated constructor stub
	}

	public ExcelFile(InputStream is) {
		try {

			book = Workbook.getWorkbook(is);
			sheet = book.getSheet(0);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// countRow = book.getSheet(0).getRows();

	}

	public ExcelFile(InputStream is, int sheetId) {
		try {
			book = Workbook.getWorkbook(is);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		this.sheetId = sheetId;
		sheet = getSheet();
	}

	public ExcelFile(File f) {
		try {

			book = Workbook.getWorkbook(f);
			sheet = book.getSheet(0);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// countRow = book.getSheet(0).getRows();

	}

	public ExcelFile(File f, int sheetId) {
		try {
			book = Workbook.getWorkbook(f);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		this.sheetId = sheetId;
		sheet = getSheet();
	}

	/**
	 * excel对象
	 */
	protected Workbook book;
	/**
	 * 总行数
	 */
	protected int countRow;
	/**
	 * 标题所占行
	 */
	protected int titleUseRow;
	/**
	 * 当前点
	 */
	protected int curRow;
	/**
	 * sheet
	 */
	protected int sheetId = 0;
	/**
	 * sheet总个数
	 */
	protected int sheetNumber;
	/**
	 * excel标题
	 */
	protected List<String> title;
	/**
	 * excel文件名
	 */
	protected String fileName;

	public String getFileName() {
		return fileName;
	}

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	/**
	 * 当前读取的sheet
	 */
	private Sheet sheet;

	private String value;

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

	public Workbook getBook() {
		return book;
	}

	public void setBook(Workbook book) {
		this.book = book;
	}

	public int getCountRow() {
		countRow = sheet.getRows();
		return countRow;
	}

	public void setCountRow(int countRow) {
		this.countRow = countRow;
	}

	public int getTitleUseRow() {
		return titleUseRow;
	}

	public void setTitleUseRow(int titleUseRow) {
		this.titleUseRow = titleUseRow;
	}

	public int getCurRow() {
		return curRow;
	}

	public void setCurRow(int curRow) {
		this.curRow = curRow;
	}

	public int getSheetId() {
		return sheetId;
	}

	public void setSheetId(int sheetId) {
		this.sheetId = sheetId;
	}

	public int getSheetNumber() {
		return sheetNumber;
	}

	public void setSheetNumber(int sheetNumber) {
		this.sheetNumber = sheetNumber;
	}

	public List<String> getTitle() {
		return title;
	}

	public void setTitle(List<String> title) {
		this.title = title;
	}

	public Sheet getSheet() {
		sheet = book.getSheet(sheetId);
		return sheet;
	}

	public void setSheet(Sheet sheet) {
		this.sheet = sheet;
	}

}
