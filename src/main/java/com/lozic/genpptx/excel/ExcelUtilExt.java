package com.lozic.genpptx.excel;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class ExcelUtilExt<T> {

	/**
	 * excle版本 2003
	 */
	private static final int EXCEL_VERSION_2003 = 1;

	/**
	 * excle版本 2007
	 */
	private static final int EXCEL_VERSION_2007 = 2;

	private String defaultDateFormat = "yyyy-MM-dd";

	public Workbook workBook;

	private Class<T> clazz;

	private List<MetaData> metaDatas;

	public List<MetaData> getMetaDatas() {
		return metaDatas;
	}

	public void setMetaDatas(List<MetaData> metaDatas) {
		this.metaDatas = metaDatas;
	}

	public int getSheetNum() {
		return sheetNum;
	}
	//返回要到导入的excel的sheet数
	public int getSheetNum(boolean isRealNumber) {
		if(isRealNumber){
			return workBook.getNumberOfSheets();
		}else{
			return sheetNum;
		}
	}

	// 设置tab页面，默认是0
	private int sheetNum = 0;

	/**
	 * 修改时间格式
	 * 
	 * @param dateFormat
	 */
	public void setDateFormat(String dateFormat) {
		this.defaultDateFormat = dateFormat;
	}

	/**
	 * 
	 * 设置读取excel的tab页面
	 * 
	 * @param sheetNum
	 */
	public void setSheetNum(int sheetNum) {
		this.sheetNum = sheetNum;
	}

	/**
	 * 
	 * Creates a new instance of ExcelUtilExt.
	 * 
	 * @param input
	 *            文件输入流
	 * @param clazz
	 *            类
	 * @param offset
	 *            数据从第几行开始获取
	 * @throws IOException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public ExcelUtilExt(InputStream input, Class<T> clazz, int offset) throws IOException, SecurityException,
            NoSuchFieldException {
		try {
			workBook = new XSSFWorkbook(input);// new XSSFWorkbook(input);
		} catch (Exception ex) {
			ex.printStackTrace();
			try {
				workBook = new HSSFWorkbook(input);
			} catch (IOException e) {
				e.printStackTrace();
				throw e;
			}

		}
		this.clazz = clazz;

	}

	/**
     * Creates a new instance of ExcelUtilExt.
     *
     * @param input     文件输入流
     * @param offset    数据从第几行开始获取
     * @throws Exception
     */
    public ExcelUtilExt(InputStream input, Class<T> clazz, int offset, int excelVersion) throws IOException {
        try {
            if (excelVersion == ExcelUtilExt.EXCEL_VERSION_2007) {
                workBook = new XSSFWorkbook(input);// new XSSFWorkbook(input);
            } else {
                workBook = new HSSFWorkbook(input);
            }
        } catch (IOException e) {
            e.printStackTrace();
            throw e;
        }

        this.clazz = clazz;

    }
	
	/**
	 * 
	 * Creates a new instance of ExcelUtilExt.
	 * 
	 * @param data
	 *            文件内容
	 * @param clazz
	 *            元数据的定义
	 * @param offset
	 *            数据从第几行开始获取
	 * @throws IOException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public ExcelUtilExt(byte[] data, Class<T> clazz, int offset) throws IOException, SecurityException,
            NoSuchFieldException {
		try {
			workBook = new XSSFWorkbook(new ByteArrayInputStream(data));// new XSSFWorkbook(input);
		} catch (Exception ex) {
			try {
				workBook = new HSSFWorkbook(new ByteArrayInputStream(data));
			} catch (IOException e) {
				throw e;
			}

		}
		this.clazz = clazz;

	}

	/**
	 * 
	 * Creates a new instance of ExcelUtilExt.
	 * 
	 * @param input
	 *            文件输入流
	 * @param metaDatas
	 *            元数据的定义
	 * @param offset
	 *            数据从第几行开始获取
	 * @throws IOException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public ExcelUtilExt(InputStream input, List<MetaData> metaDatas, Class<T> clazz, int offset, List<T> result,
                        boolean readMetaData) throws IOException, SecurityException, NoSuchFieldException {
		try {
			workBook = new XSSFWorkbook(input);// new XSSFWorkbook(input);
		} catch (Exception ex) {
			ex.printStackTrace();
			try {
				workBook = new HSSFWorkbook(input);
			} catch (IOException e) {
				e.printStackTrace();
				throw e;
			}

		}
		this.clazz = clazz;
		this.metaDatas = metaDatas;
		try {
			// if(readMetaData)
			// {
			// parseHead(offset, sheetNum,metaDatas);
			// }
			parseContent(offset, sheetNum, result);
		} catch (InstantiationException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 
	 * getTitleCount:得到excel标题数据信息. <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 * 
	 * @return
	 */
	public int getTitleCount(int offset) {
		Sheet sheet = workBook.getSheetAt(sheetNum);
		Row row = sheet.getRow(offset);// 得到标题行
		if (null == row) {
			return 0;
		} else {
			return row.getLastCellNum();
		}
	}

	/**
	 * 
	 * 解析head
	 * 
	 * @param offset
	 * @param sheetNum
	 * @param metaDatas
	 */
	public List<ErrorMsgVO> parseHead(int offset, int sheetNum, List<MetaData> metaDatas, List<ErrorMsgVO> list) {
		Sheet sheet = workBook.getSheetAt(sheetNum);


			Row row = sheet.getRow(offset);

			for (int cellNum = 0, length = metaDatas.size(); cellNum < length; cellNum++) {
				Cell cell = row.getCell(cellNum);
				MetaData metaData = metaDatas.get(cellNum);
                // getRichStringCellValue方法只支持这几种格式获取RichString，其他格式会报异常 - 刘杰
                boolean canReadRichStringCellValue = cell.getCellType() == Cell.CELL_TYPE_STRING ||
                        cell.getCellType() == Cell.CELL_TYPE_FORMULA ||
                        cell.getCellType() == Cell.CELL_TYPE_BLANK;

                metaData.setComment((canReadRichStringCellValue && cell.getRichStringCellValue()!=null)
                        ? cell.getRichStringCellValue().getString()
                        : "");
				if (!metaData.compareField()) {
					list.add(new ErrorMsgVO("第1行第" + (cellNum + 1) + "列", "该列标题不正确", "该列名称为"
							+ metaData.getDefaultComments()));
				} else {

				}
			}


		return list;

	}


	/**
	 *
	 * 制定行获取head
	 *
	 * @param offset
	 * @param sheetNum
	 * @param metaDatas
	 */
	public List<ErrorMsgVO> parseHeadFromN(int offset, int sheetNum, List<MetaData> metaDatas, List<ErrorMsgVO> list) {
		Sheet sheet = workBook.getSheetAt(sheetNum);

		for (int rowNum = 0; rowNum < offset; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				continue;
			}

			for (int cellNum = 0, length = metaDatas.size(); cellNum < length; cellNum++) {
				Cell cell = row.getCell(cellNum);
				MetaData metaData = metaDatas.get(cellNum);
                // getRichStringCellValue方法只支持这几种格式获取RichString，其他格式会报异常 - 刘杰
                boolean canReadRichStringCellValue = cell.getCellType() == Cell.CELL_TYPE_STRING ||
                        cell.getCellType() == Cell.CELL_TYPE_FORMULA ||
                        cell.getCellType() == Cell.CELL_TYPE_BLANK;

                metaData.setComment((canReadRichStringCellValue && cell.getRichStringCellValue()!=null)
                        ? cell.getRichStringCellValue().getString()
                        : "");
				if (!metaData.compareField()) {
					list.add(new ErrorMsgVO("第1行第" + (cellNum + 1) + "列", "该列标题不正确", "该列名称为"
							+ metaData.getDefaultComments()));
				} else {
					continue;
				}
			}

		}
		return list;

	}

	public int getLastValidRowNum(int offset,Sheet sheet) {
		for (int rowNum = offset, rowLength = sheet.getLastRowNum(); rowNum <= rowLength; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				continue;
			}
			Cell cell = row.getCell(0);
			cell.setCellType(Cell.CELL_TYPE_STRING);
			if (null == cell.getStringCellValue() || "".equals(cell.getStringCellValue())) {
				return rowNum;
			}
		}
		return offset;
	}

	/**
	 * parse:解析数据，从指定的行开始. <br/>
	 * 
	 * @param offset
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public void parseContent(int offset, int sheetNum, List<T> result) throws InstantiationException,
            IllegalAccessException, SecurityException, NoSuchFieldException {
		Sheet sheet = workBook.getSheetAt(sheetNum);
		if (sheet == null) {
			throw new RuntimeException("sheetNum " + sheetNum + " doesn't exist!");
		}

		for (int rowNum = offset, rowLength = getLastValidRowNum(offset,sheet)-1 ; rowNum <= rowLength; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				continue;
			}

			T instance = clazz.newInstance();
			result.add(instance);

			for (int cellNum = 0, length = metaDatas.size(); cellNum < length; cellNum++) {
				Cell cell = row.getCell(cellNum);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				MetaData metaData = metaDatas.get(cellNum);
				getValue(cell, metaData, instance);
			}

		}

	}
	
	
	/**
	 * parse:采购品维护模块 导入采购品  解析excel <br/>
	 * 
	 * @param offset
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public void parseContent4Import(int offset, int sheetNum, List<T> result) throws InstantiationException,
            IllegalAccessException, SecurityException, NoSuchFieldException {
		Sheet sheet = workBook.getSheetAt(sheetNum);
		if (sheet == null) {
			throw new RuntimeException("sheetNum " + sheetNum + " doesn't exist!");
		}

		for (int rowNum = offset, rowLength = sheet.getLastRowNum(); rowNum <= rowLength; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				continue;
			}

			T instance = clazz.newInstance();
			result.add(instance);

			for (int cellNum = 0, length = metaDatas.size(); cellNum < length; cellNum++) {
				Cell cell = row.getCell(cellNum);
				MetaData metaData = metaDatas.get(cellNum);
				getValue4Import(cell, metaData, instance);
			}

		}

	}
	
	
	/**
	 * @desc 采购品维护模块 导入采购品 获取值 全部按照String类型读取
	 * @param cell
	 *            单元
	 * @param metaData
	 *            源数据
	 * @return
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	private void getValue4Import(Cell cell, MetaData metaData, Object instance) throws SecurityException, NoSuchFieldException {
 		String setMethodName = "set" + toFirstLetterUpperCase(metaData.getField());
 		if(null != cell){
 			cell.setCellType(Cell.CELL_TYPE_STRING);
			String value=cell.getStringCellValue();
			setValue(instance, setMethodName, value);
		}else{
			setValue(instance, setMethodName,"");
		}

	}

	/**
	 * @desc 获取值，如果包含中午进行全角转换未做
	 * @param cell
	 *            单元
	 * @param metaData
	 *            源数据
	 * @return
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	private void getValue(Cell cell, MetaData metaData, Object instance) throws SecurityException, NoSuchFieldException {

		String setMethodName = "set" + toFirstLetterUpperCase(metaData.getField());

		// instance.getClass().getField(metaData.getField()).getType();
		Class<?> declareClass = instance.getClass().getDeclaredField(metaData.getField()).getType();
		if(null != cell){
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BOOLEAN:
				setValue(instance, setMethodName, cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					Date date = cell.getDateCellValue();
					setValue(instance, setMethodName, date);
				} else {
					if (declareClass.equals(java.lang.String.class)) {
						cell.setCellType(Cell.CELL_TYPE_STRING);
						String cellvalue =StringUtils.trimToEmpty(cell.getRichStringCellValue().getString());
						if (cellvalue.endsWith(".0")) {
							cellvalue = cellvalue.replace(".0", "");
						}
						setValue(instance, setMethodName, replaceString(cellvalue));
					} else {
						String cellvalue = String.valueOf(cell.getNumericCellValue()).trim();
						if (cellvalue.endsWith(".0")) {
							cellvalue = cellvalue.replace(".0", "");
						}
						setNumberValue(instance, declareClass, setMethodName, cellvalue);

					}
				}
				break;

			case Cell.CELL_TYPE_STRING:
				String cellvalue = StringUtils.trimToEmpty(cell.getRichStringCellValue().getString());
				if (declareClass.getSuperclass().equals(java.lang.Number.class)) {
					setNumberValue(instance, declareClass, setMethodName, cellvalue);
				} else if (declareClass.equals(Date.class))// 日期
				{
					DateFormat df = new SimpleDateFormat(defaultDateFormat);
					try {
						setValue(instance, setMethodName, df.parse(cellvalue));
					} catch (ParseException e) {
						e.printStackTrace();
					}
	
				} else {
					setValue(instance, setMethodName, replaceString(cellvalue));
				}
	
				break;
	
			}
		}else{
			setValue(instance, setMethodName,"");
		}

	}

	public Workbook getWorkBook() {
		return workBook;
	}

	public void setWorkBook(Workbook workBook) {
		this.workBook = workBook;
	}

	/**
	 * 数字类型，调用 valuelof方法
	 * 
	 * @param instance
	 * @param declareClass
	 * @param cellvalue
	 */
	private void setNumberValue(Object instance, Class<?> declareClass, String setMethodName, String cellvalue) {
		try {
			Method staticMethod = declareClass.getMethod("valueOf", new Class[] { String.class });
			Object value = staticMethod.invoke(declareClass, cellvalue);
			setValue(instance, setMethodName, value);
		} catch (SecurityException e) {
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 设置值 setValue:(这里用一句话描述这个方法的作用). <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 * 
	 * @param setMethodName
	 * @param instance
	 * @param value
	 */
	private void setValue(Object instance, String setMethodName, Object value) {
		try {
			instance.getClass().getMethod(setMethodName, value.getClass()).invoke(instance, value);
		} catch (IllegalArgumentException e) {
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 首字母大写
	 * 
	 * @param str
	 * @return
	 */
	private String toFirstLetterUpperCase(String str) {
		if (str == null || str.length() < 2) {
			return str;
		}
		String firstLetter = str.substring(0, 1).toUpperCase();
		return firstLetter + str.substring(1, str.length());
	}

	private static String replaceString(String str){
		if(StringUtils.isEmpty(str)){
			return "";
		}
		return str.replaceAll("[\\n\\r]", "");

	}
}
