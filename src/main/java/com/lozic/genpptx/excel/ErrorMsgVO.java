/**
 * Project Name:bpo-model
 * File Name:PlanVO.java
 * Package Name:cn.bidlink.bpo.model.domain
 * Date:2013-5-6上午11:29:30
 * Author  <a href="mailto:changzhiyuan@ebnew.com">changzhiyuan</a>
 * Copyright (c) 2013
 *
 */
package com.lozic.genpptx.excel;

/**
 * ClassName: PlanVO <br/>
 * Function: TODO 错误信息对话框的提示. <br/>
 * Reason: TODO ADD REASON(可选). <br/>
 * date: 2013-5-6 上午11:29:30 <br/>
 *
 * @author Administrator
 * @version 
 * @since JDK 1.6
 */
public class ErrorMsgVO {

    /**
    * @Fields serialVersionUID : TODO(用一句话描述这个变量表示什么)
    */
    private static final long serialVersionUID = -2391161639116390382L;

    /**
     * 错误所在行列值
     */
    private String location;

    /**
     * 错误造成原因
     */
    private String errorMsg;

    /**
     * 错误修改意见
     */
    private String correctMsg;

    public ErrorMsgVO() {
        super();
    }

    public ErrorMsgVO(String location, String errorMsg, String correctMsg) {
        super();
        this.location = location;
        this.errorMsg = errorMsg;
        this.correctMsg = correctMsg;
    }

    public String getLocation() {
        return location;
    }

    public void setLocation(String location) {
        this.location = location;
    }

    public String getErrorMsg() {
        return errorMsg;
    }

    public void setErrorMsg(String errorMsg) {
        this.errorMsg = errorMsg;
    }

    public String getCorrectMsg() {
        return correctMsg;
    }

    public void setCorrectMsg(String correctMsg) {
        this.correctMsg = correctMsg;
    }

}
