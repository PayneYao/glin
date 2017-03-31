package com.lozic.genpptx;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by wangqingwu on 16/11/22.
 * Since 16/11/22
 * Author Simon Gaius
 */
public class SlideEntity {

    private List<ElementBean> elementBeanList = null;

    public SlideEntity(){
        elementBeanList = new ArrayList<ElementBean>();
    }

    public List<ElementBean> addElement(ElementBean elementBean){
        if(elementBeanList!=null&&elementBean!=null) {
            elementBeanList.add(elementBean);
        }
        return elementBeanList;
    }

    public List<ElementBean> getElementBeanList(){
        return elementBeanList;
    }


}
