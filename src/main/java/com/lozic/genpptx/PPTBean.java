package com.lozic.genpptx;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by wangqingwu on 16/11/22.
 * Since 16/11/22
 * Author Simon Gaius
 */
public class PPTBean {
    private String name = null;

    private List<SlideEntity> slideEntitiesList = null;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCoverPath() {
        return coverPath;
    }

    public void setCoverPath(String coverPath) {
        this.coverPath = coverPath;
    }



    public List<SlideEntity> addSlide(SlideEntity slideEntity){

        if(slideEntitiesList!=null&& slideEntity!=null) {
            slideEntitiesList.add(slideEntity);
        }
        return slideEntitiesList;

    }


    private String coverPath = null;

    public List<SlideEntity> getSlideEntitiesList(){
        return slideEntitiesList;
    }

    public PPTBean() {
        slideEntitiesList = new ArrayList<SlideEntity>();
    }

}
