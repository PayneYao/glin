package com.lozic.genpptx.excel;

import org.apache.commons.collections.CollectionUtils;

import java.util.List;

/**
 *  shiyongren 采购品导入结果对象
 *
 * @version : Ver 1.0
 * @author	: <a href="mailto:shiyongren@ebnew.com">shiyongren</a>
 * @date	: 2015年9月16日 上午9:50:29
 */
public class ExcelOperationResltVo {
	private boolean importRes; //导入结果
	private String msg; //返回信息
	private List<ErrorMsgVO> errorMsgs; //验证的错误信息
	public boolean isImportRes() {
		return importRes;
	}
	public void setImportRes(boolean importRes) {
		this.importRes = importRes;
	}
	public String getMsg() {
		return msg;
	}
	public void setMsg(String msg) {
		this.msg = msg;
	}
	public List<ErrorMsgVO> getErrorMsgs() {
		return errorMsgs;
	}
	public void setErrorMsgs(List<ErrorMsgVO> errorMsgs) {
		this.errorMsgs = errorMsgs;
	}
	public boolean getHasErrorMsgs(){
		return CollectionUtils.isNotEmpty(errorMsgs);
	}
	
}
