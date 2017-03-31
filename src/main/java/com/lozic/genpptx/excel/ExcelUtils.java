package com.lozic.genpptx.excel;

import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang.StringUtils;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * ClassName: ExcelUtils <br/>
 * Function: TODO excel处理类. <br/>
 * Reason: TODO ADD REASON(可选). <br/>
 * date: 2013-5-6 下午3:21:56 <br/>
 *
 * @author yingboshan
 * @version
 * @since JDK 1.6
 */
public class ExcelUtils {

	private static ExcelUtils eu = new ExcelUtils();

	private ExcelUtils() {
	};

	public static ExcelUtils getInstance() {
		return eu;
	}

	public int getColumnCount(Sheet sheet) {
		return sheet.getColumns();
	}

	/**
	 *
	 * validateTitle:验证title是否合法. <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 *
	 * @param msg
	 *            标题行不匹配
	 *
	 * @return success:成功
	 */
	public List<ErrorMsgVO> validateTitle(Sheet sheet, int row,
                                          List<ErrorMsgVO> list, String msg, String[] s) {
		Cell[] cells = sheet.getRow(row);

		for (int i = 0; i < cells.length; i++) {
			if (cells[i].getContents().equals(s[i])) {
				continue;
			} else {
				list.add(new ErrorMsgVO(changePosInfo(0, i), msg, "该列名称为"
						+ s[i]));
			}
		}
		return list;

	}

	/**
	 *
	 * validateTypeColumn:验证某列是否都为数字类型. <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 *
	 * @param col
	 *            待验证列序号
	 */
	public List<ErrorMsgVO> validateTypeColumn(Sheet sheet, int col,
                                               List<ErrorMsgVO> list, CellType validateType, String correctMsg,
                                               boolean isNull, String errorMsg) {
		Cell[] cells = sheet.getColumn(col);
		for (int i = 1; i < cells.length; i++) {
			if (cells[i].getType() != validateType
					&& isNullRow(sheet.getRow(i))) {
				if (isNull && cells[i].getContents().isEmpty())
					continue;
				list.add(new ErrorMsgVO(changePosInfo(cells[i].getRow(),
						cells[i].getColumn()), errorMsg, correctMsg));
			}
		}
		return list;
	}

	/**
	 *
	 * validateTypeColumn:验证某列是否都为空串. <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 *
	 * @param col
	 * @param isNull
	 *            是否为空 待验证列序号
	 */
	public List<ErrorMsgVO> validateTypeColumn(Sheet sheet, int col,
                                               List<ErrorMsgVO> list, boolean isNull, String correctMsg,
                                               String errorMsg) {
		Cell[] cells = sheet.getColumn(col);
		for (int i = 1; i < cells.length; i++) {

			if (!isNull && cells[i].getContents().isEmpty()
					&& isNullRow(sheet.getRow(i))) {
				list.add(new ErrorMsgVO(changePosInfo(cells[i].getRow(),
						cells[i].getColumn()), errorMsg, correctMsg));

			}

		}
		return list;
	}

	/**
	 *
	 * changePosInfo:转换位置信息变成某行某列. <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 *
	 * @param row
	 * @param col
	 * @return
	 */
	private String changePosInfo(int row, int col) {
		if (row >= 0 && col >= 0) {
			return "第" + (row + 1) + "行第" + (col + 1) + "列";
		}
		return "参数不能为空";
	}

	/**
	 *
	 * validateFileExt:验证文件后缀名. <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 *
	 * @param fileName
	 * @return
	 */
	public boolean validateFileExt(String fileName) {
		if (fileName.endsWith(".xls")) {
			return true;
		}
		return false;
	}

	/**
	 *
	 * queryByRowId:根据行号查询指定行<br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 *
	 * @return
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 */
	public Object queryByRowId(int rowid, Sheet sheet, String[] properties,
                               Class clazz) {
		Object o = null;
		try {
			o = clazz.newInstance();

			Cell[] cells = sheet.getRow(rowid);
			for (int i = 0; i < cells.length; i++) {
				// 对于日期格式的单元格需另行处理，日期处理类待扩展
				if (CellType.DATE.equals(cells[i].getType())) {
					try {
						SimpleDateFormat dateFormat = new SimpleDateFormat(
								"yyyy-MM-dd");
						BeanUtils.copyProperty(o, properties[i],
								dateFormat.format((cells[i].getContents())));
					} catch (IllegalAccessException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (InvocationTargetException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					continue;

				}
				try {
					if (properties[i].equals("needDate")
							&& cells[i].getContents().equals("")) {

					} else {
						BeanUtils.copyProperty(o, properties[i],
								cells[i].getContents());
					}
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

			}

		} catch (InstantiationException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IllegalAccessException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		return o;

	}

	public List queryAll(Sheet sheet, int titleUse, String[] properties,
                         Class clazz) {
		List list = new ArrayList();
		StringBuilder sb = new StringBuilder();
		// 获得结束行
		/*
		 * int totalCount = 0; while(totalCount < sheet.getRows()){
		 * sheet.getCell(0, totalCount).getContents().trim().e; totalCount++; }
		 */
		// 从数据行开始读取数据
		for (int i = titleUse; i < sheet.getRows(); i++) {
			System.out.println(isNullRow(sheet.getRow(i)));
			if (isNullRow(sheet.getRow(i))) {
				list.add(queryByRowId(i, sheet, properties, clazz));
			}

		}
		return list;
	}

	/**
	 *
	 * isNullRow:是否空行. <br/>
	 * TODO(这里描述这个方法适用条件 – 可选).<br/>
	 *
	 * @param true:非空行
	 * @return
	 */
	public boolean isNullRow(Cell[] c) {
		for (Cell cell : c) {

			if (StringUtils.trimToEmpty(cell.getContents()).equals("")) {
				continue;
			} else {
				return true;
			}
		}
		return false;
	}

	/**
	 *
	 * createSheetByList:此方法目前仅用于采购计划创建Excel Sheet. <br/>
	 *
	 * @param fileName
	 * @param title
	 * @param planItems
	 * @return
	 */
	public File createSheetByList(String fileName, List<String> title, List<String[]> planItems) throws IOException, WriteException {

		File f = new File(fileName);
		ExcelFile ef = new ExcelFile(f);
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

		WritableWorkbook wb = null;
		try {
			// 创建可写入的 Excel工作薄
			wb = ef.book.createWorkbook(f);
			// 创建Excel工作表
			WritableSheet ws = wb.createSheet("sheet1", 0);

			/** 插入标题行 */
			Label label1 = null;
			for (int i = 0; i < title.size(); i++) {
				label1 = new Label(i, 0, title.get(i));
				ws.addCell(label1);
			}

			/** 插入数据区 */
			int colNum = 0;
			for (int i = 0; i < planItems.size(); i++) {
				if (title.size() > 10 && planItems.get(i).length > 10) {// 有编码
					colNum = 0;
					ws.addCell(new Label(colNum, i + 1, planItems.get(i)[10]));
					colNum = colNum + 1;
				}
				ws.addCell(new Label(0 + colNum, i + 1, planItems.get(i)[0]));
				ws.addCell(new Label(1 + colNum, i + 1, planItems.get(i)[1]));
				ws.addCell(new Label(2 + colNum, i + 1, planItems.get(i)[2]));
				ws.addCell(new Label(3 + colNum, i + 1, planItems.get(i)[3]));
				ws.addCell(new Label(4 + colNum, i + 1, planItems.get(i)[4]));
				ws.addCell(new Label(5 + colNum, i + 1, planItems.get(i)[5]));
				ws.addCell(new Label(6 + colNum, i + 1, planItems.get(i)[6]));
				ws.addCell(new Label(7 + colNum, i + 1, planItems.get(i)[7]));
				ws.addCell(new Label(8 + colNum, i + 1, sdf.format(planItems.get(i)[8])));
				ws.addCell(new Label(9 + colNum, i + 1, planItems.get(i)[9]));
			}

			// 添加Label对象

			// ws.addCell(label1);

			// 添加带有字型Formatting的对象
			/*
			 * WritableFont wf = new WritableFont( WritableFont.TIMES, 18,
			 * WritableFont.BOLD, true); WritableCellFormat wcfF = new
			 * WritableCellFormat( wf); Label labelCF = new Label(1, 0,
			 * "This is a Label Cell", wcfF); ws.addCell(labelCF);
			 *
			 * // 添加带有字体颜色Formatting的对象 WritableFont wfc = new WritableFont(
			 * WritableFont.ARIAL, 10, WritableFont.NO_BOLD, false,
			 * UnderlineStyle.NO_UNDERLINE, jxl.format.Colour.RED);
			 * WritableCellFormat wcfFC = new WritableCellFormat( wfc); Label
			 * labelCFC = new Label(1, 0, "This is a Label Cell", wcfFC);
			 * ws.addCell(labelCFC);
			 *
			 * // 添加Number对象 jxl.write.Number labelN = new jxl.write.Number(0,
			 * 1, 3.1415926); ws.addCell(labelN);
			 *
			 * // 添加带有formatting的 Number对象 NumberFormat nf = new
			 * NumberFormat("#.##"); WritableCellFormat wcfN = new
			 * WritableCellFormat( nf); jxl.write.Number labelNF = new
			 * jxl.write.Number(1, 1, 3.1415926, wcfN); ws.addCell(labelNF);
			 *
			 * // 添加Boolean对象 jxl.write.Boolean labelB = new
			 * jxl.write.Boolean(0, 2, false); ws.addCell(labelB);
			 *
			 * // 添加DateTime对象 DateTime labelDT = new DateTime(0, 3, new
			 * java.util.Date()); ws.addCell(labelDT);
			 *
			 * // 添加带有formatting的DateFormat对象 DateFormat df = new DateFormat(
			 * "dd MM yyyy hh:mm:ss"); WritableCellFormat wcfDF = new
			 * WritableCellFormat( df); DateTime labelDTF = new DateTime(1, 3,
			 * new java.util.Date(), wcfDF); ws.addCell(labelDTF);
			 */

			wb.write(); // 写入Exel工作表
		} finally {
			if(wb != null){
                try {
                    wb.close();// 关闭Excel工作薄对象
                } catch (IOException e) {
                    e.printStackTrace();
                } catch (WriteException e) {
                    e.printStackTrace();
                }
            }
		}

		return f;
	}

}
