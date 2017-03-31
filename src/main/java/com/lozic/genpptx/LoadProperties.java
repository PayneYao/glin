package com.lozic.genpptx;

import com.lozic.genpptx.util.JProperties;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Properties;

/**
 * Created by wangqingwu on 16/11/22.
 * Since 16/11/22
 * Author Simon Gaius
 */
public class LoadProperties {

    public List<String> getAllLines(String filePath) {

        List<String> lines = null;
        try {
            lines = Files.readAllLines(Paths.get(filePath), StandardCharsets.UTF_8);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return lines;

    }

    public PPTBean getPPTBean(Properties p,String rootPath) throws IOException {
        PPTBean pptBean = null;
          Path confPath = Paths.get("/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/conf.properties");
    // private int titleCount = 16;
     String project = "test";
//          Properties p = new Properties();
  //      p = JProperties.loadProperties(confPath.toString(), JProperties.BY_PROPERTIES);
        project = p.getProperty("excel.project");
        //String rootPath = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
        String prjName=project;
        String confFilePath = Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString();
        //String confFile = "/Users/wangqingwu/Projects/gen-pptx/ppcloud.properties";

        List<String> lines = getAllLines(confFilePath);

        SlideEntity slideEntity = null;

        for (String line : lines) {
            ElementBean eb = null;
            if ("#".equals(line)) {
                pptBean = new PPTBean();
                continue;
            }
            if ("##".equals(line) && pptBean != null) {
                if (slideEntity != null) {
                    pptBean.addSlide(slideEntity);
                }
                slideEntity = new SlideEntity();
                continue;
            }
            if ("###".equals(line) && pptBean != null) {
                pptBean.addSlide(slideEntity);
                return pptBean;

            }
            if (line != null && slideEntity != null) {
                eb = new ElementBean();
                String[] ll = line.split(";");
                if (ll != null && ll.length == 3) {
                    String[] pos = ll[0].split(",");
                    if (pos != null && pos.length == 4) {
                        eb.setX(Double.parseDouble(pos[0].trim()));
                        eb.setY(Double.parseDouble(pos[1].trim()));
                        eb.setWidth(Double.parseDouble(pos[2].trim()));
                        eb.setHeight(Double.parseDouble(pos[3].trim()));
                        eb.setType(Integer.parseInt(ll[1].trim()));
                        eb.setContent(ll[2].trim());
                        slideEntity.addElement(eb);
                    }
                }
                /**
                if (slideEntity != null) {
                    slideEntity.addElement(eb);
                }
                 */
            }

        }


        return pptBean;
    }

    public LoadProperties() {
    }
}
