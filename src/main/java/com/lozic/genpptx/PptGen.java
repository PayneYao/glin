package com.lozic.genpptx;

import java.nio.charset.StandardCharsets;

import com.lozic.genpptx.excel.ErrorMsgVO;
import com.lozic.genpptx.excel.ExcelUtilExt;
import com.lozic.genpptx.excel.MetaData;
import com.lozic.genpptx.model.ProjectDetail;
import com.lozic.genpptx.util.JProperties;
import org.apache.poi.hslf.blip.PICT;
import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.sl.usermodel.*;
import org.apache.poi.ss.usermodel.FontCharset;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;

import javax.swing.plaf.ColorUIResource;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.*;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * Created by wangqingwu on 16/11/17.
 * Since 16/11/17
 * Author Simon Gaius
 */
public class PptGen {
  private Path templatePath;
  private Path outputPath;
  private Path path;
  private Path out5;
  private Path tp1;
  private Path excelPath;
  private int offset = 2;
  private int titleCount = 16;
  private Path rootPath = Paths.get("d:\\gaius\\projects\\");
  private String prjName = "KFC-北京";
  private int getTitleCountBrief = 4;
  private Properties p = new Properties();
  private String templateFileExt = ".pptx";
  private String defaultTemplateFileName = "template1.pptx";
  //simon1private Path confPath = Paths.get("pptgenerator/projects/conf.properties");
  // private int titleCount = 16;
  //  private String project = "test.xlsx";

  private int offsetBrief = 2;

  public PptGen(Properties p, String root, String projectName) {
    this.p = p;
    String model = p.getProperty("excel.model");
    String v = p.getProperty("excel." + model + ".header.count");
    titleCount = Integer.parseInt(v);
    this.rootPath = Paths.get(root);
    this.prjName = projectName;
    //project = p.getProperty("excel.project");
    templatePath = Paths.get(root, projectName, "cover", defaultTemplateFileName);
  }

  public void init() throws URISyntaxException, IOException {
    // templatePath = Paths.get(this.getClass().getResource("test2.pptx").toURI());
    Properties p = new Properties();
    //simon1p = JProperties.loadProperties(confPath.toString(), JProperties.BY_PROPERTIES);

    //prjName = project;
    outputPath = Paths.get("target/out2.pptx");
    excelPath = Paths.get("1.xlsx");
    out5 = Paths.get(this.getClass().getResource("out5.pptx").toURI());
    tp1 = Paths.get(this.getClass().getResource("tp1.pptx").toURI());
    //  Files.deleteIfExists(outputPath);
    path = Paths.get(this.getClass().getResource("signals.png").toURI());
    templatePath = Paths.get(rootPath.toString(), prjName, "cover", defaultTemplateFileName);
    String key = "";

    String model = p.getProperty("excel.model");
    String v = p.getProperty("excel." + model + ".header.count");
    titleCount = Integer.parseInt(v);
    System.out.println("excel.2.header.count:" + v);

  }

  public void prepareDirectory(String root, String prjname)
      throws IOException, NoSuchFieldException {
    String f = root + prjname + ".xlsx";
    String fdir = Paths.get(root, prjname).toString();
    //getFileNameNoEx(fo.getName());
    File file = new File(fdir);
    if (!file.exists() || !file.isDirectory()) {
      file.mkdirs();
    }
    File f1 = new File(root + prjname + "/cover");
    if (!f1.exists()) {
      f1.mkdirs();
    }
    File ff = new File(root + prjname + "/conf");
    if (!ff.exists()) {
      ff.mkdirs();
    }

    File f2 = new File(root + prjname + "/out");
    if (!f2.exists()) {
      f2.mkdirs();
    }
    List<ProjectDetail> result = getCommusByPrj(f);
    for (int i = 0; i < result.size(); i++) {
      ProjectDetail pd = result.get(i);
      File f3 = new File(root + prjname + "/" + pd.getName());
      if (!f3.exists()) {
        f3.mkdirs();
      }
      File f5 = new File(root + prjname + "/" + pd.getName() + "/小区");
      File f6 = new File(root + prjname + "/" + pd.getName() + "/门栋");
      File f7 = new File(root + prjname + "/" + pd.getName() + "/广告");

      if (!f5.exists()) {
        f5.mkdirs();

      }
      if (!f6.exists()) {
        f6.mkdirs();
      }
      if (!f7.exists()) {
        f7.mkdirs();
      }
    }

  }

  //生成属性文件
  //@Test
  public void createConfig3(Properties p, String rootPath, String prjName)
      throws IOException, NoSuchFieldException {
    File file = new File(
        Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString());
    //if(!file.exists()) {
    //            file.createNewFile();
    //        }
    //String regionName = "金都杭城商务楼";
    FileOutputStream fos = new FileOutputStream(file);

    //        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
    BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));

    String f = Paths.get(rootPath, prjName + ".xlsx").toString();

    //String root = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
    File fo = new File(f);
    String prjname = getFileNameNoEx(fo.getName());
    String pstr = "#";
    bw.write(pstr);
    bw.newLine();
    List<ProjectDetail> result = getCommusByPrj(f);
    for (ProjectDetail pd : result) {
      String[] picList = getPicArray(pd.getName());
      List<String> commuPicList = getCommuPicList(pd.getName());
      int pageSize = picList.length / 2;
      for (int i = 0; i < pageSize; i++) {
        String line0 = "##";
        bw.write(line0);
        bw.newLine();
        String line1 =
            p.getProperty("pos.text.line1", "70, 50, 450, 300;3;") + "城市##广告发布日期##监测日期##媒体位置!!!" +
                pd.getRegion() + "##" + pd.getDate1() + "##" + pd.getDate2() + "##" + pd.getName()
                + "/" + pd.getPositionDesc();
        String line4 = p.getProperty("pos.pic.line2", "80, 300, 150, 150;2;") + picList[i * 2];
        String line5 = p.getProperty("pos.pic.line3", "350, 300, 150, 150;2;") + picList[i * 2 + 1];
        bw.write(line1);
        bw.newLine();
        bw.write(line4);
        bw.newLine();
        bw.write(line5);
        bw.newLine();
      }
    }
    String pendstr = "###";
    bw.write(pendstr);
    bw.newLine();
    bw.close();
    fos.close();
  }

  //生成属性文件
  // @Test
  public void createConfig2(Properties p, String rootPath, String prjName)
      throws IOException, NoSuchFieldException {
    //广本北京
    File file = new File(
        Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString());
    //if(!file.exists()) {
    //            file.createNewFile();
    //        }
    //String regionName = "金都杭城商务楼";
    FileOutputStream fos = new FileOutputStream(file);

    //BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
    BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));

    String f = Paths.get(rootPath, prjName + ".xlsx").toString();

    String root = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
    File fo = new File(f);
    String prjname = getFileNameNoEx(fo.getName());
    String pstr = "#";
    bw.write(pstr);
    bw.newLine();
    List<ProjectDetail> result = getCommusByPrj(f);
    for (ProjectDetail pd : result) {
      String[] picList = getPicArray(pd.getName());
      List<String> commuPicList = getCommuPicList(pd.getName());
      //int pageSize = picList.length / 2;
        int pageSize=picList.length;
     // int pageSize = (int) Math.ceil((picList.length + 1) / 2.0);
      for (int i = 0; i < pageSize; i++) {
        String line0 = "##";
        bw.write(line0);
        bw.newLine();

        String line1 =
            p.getProperty("pos.text.line1", "70, 50, 450, 300;3;") + "城市##广告发布日期##监测日期##媒体位置!!!" +
                pd.getRegion() + "##" + pd.getDate1() + "##" + pd.getDate2() + "##" + pd.getName()
                + "/" + pd.getPositionDesc();
     //   String line4="";
     //   String line5="";
     //   if(picList.length%2==1){
         // if(i==0) {
        String  line4 = p.getProperty("pos.pic.line2", "80, 300, 150, 150;2;") + commuPicList.get(0);
         String    line5 = p.getProperty("pos.pic.line3", "350, 300, 150, 150;2;") + picList[i];
     //     }else{
       //     line4 = p.getProperty("pos.pic.line2", "80, 300, 150, 150;2;") + picList[i*2-1];
      //      if(i*2<=(picList.length-1)) {
      //        line5 = p.getProperty("pos.pic.line3", "350, 300, 150, 150;2;") + picList[i * 2];
         //   }
       //   }
      //  }else if(picList.length%2==0){
        /**
          if(i==0){
            line4 = p.getProperty("pos.pic.line2", "80, 300, 150, 150;2;") + commuPicList.get(0);
            line5 = p.getProperty("pos.pic.line3", "350, 300, 150, 150;2;") + picList[0];
          }else{
            line4 = p.getProperty("pos.pic.line2", "80, 300, 150, 150;2;") + picList[i*2-1];
            line5 = p.getProperty("pos.pic.line3", "350, 300, 150, 150;2;") + picList[0];
          }
        }
         */

     //   String line4 = p.getProperty("pos.pic.line2", "80, 300, 150, 150;2;") + picList[i * 2];
       // String line5 = p.getProperty("pos.pic.line3", "350, 300, 150, 150;2;") + picList[i * 2 + 1];
        bw.write(line1);
        bw.newLine();
        bw.write(line4);
        bw.newLine();
    //    if(!"".equals(line5)) {
          bw.write(line5);
          bw.newLine();
      //  }
      }
    }
    String pendstr = "###";
    bw.write(pendstr);
    bw.newLine();
    bw.close();
    fos.close();
  }

  /**
   * @throws IOException
   * @throws NoSuchFieldException
   */

  // @Test
  public void createConfig5(Properties p, String rootPath, String prjName)
      throws IOException, NoSuchFieldException {

   /*
    客户名称：搜狗
发布城市：北京市
发布内容：搜狗
发布位置：安慧里公寓A/朝阳区安慧桥东北
合约数量：2                        实际上画数：2
发布时间：16.11.12-16.12.9        报告完成：16.11.15

     */

    File file = new File(
        Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString());
    //if(!file.exists()) {
    //            file.createNewFile();
    //        }
    //String regionName = "金都杭城商务楼";
    FileOutputStream fos = new FileOutputStream(file);

    //        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
    BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));
    String f = Paths.get(rootPath, prjName + ".xlsx").toString();

    //String root = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
    String fontSize = ":18";
    File fo = new File(f);
    String prjname = getFileNameNoEx(fo.getName());
    String pstr = "#";
    bw.write(pstr);
    bw.newLine();
    List<ProjectDetail> result = getCommusByPrj(f);
    for (ProjectDetail pd : result) {
      String[] picList = getPicArray(pd.getName());
      List<String> commuPicList = getCommuPicList(pd.getName());
      int picCount = picList.length;

      int pageSize = picList.length;
      for (int i = 0; i < pageSize; i++) {
        String line0 = "##";
        bw.write(line0);
        bw.newLine();
        String space = "";
        if((Integer.parseInt(pd.getContractCount())<100)){
          space =" ";
        }
        if(Integer.parseInt(pd.getContractCount())<10){
            space="  ";
        }
        String line3 =
            p.getProperty("pos.text.line1", "85, 50, 600, 200;1;") + "客户名称：" + prjName + p.getProperty("font.line1") +
                "###" + "发布城市：" + pd.getRegion() + p.getProperty("font.line1")  +
                "###" + "发布内容：" + prjName + p.getProperty("font.line1")  +
                "###" + "发布位置：" + pd.getName() + "/" + pd.getPositionDesc() + p.getProperty("font.line1")  +
                "###" + "合约数量：" + pd.getContractCount()+ space + "                                实际上画数：" + pd
                .getRealCount() + p.getProperty("font.line1")  +
                "###" + "发布时间：" + pd.getDate1() + "                   报告完成：" + pd.getDate2()
                + p.getProperty("font.line1") ;

        String line4 = p.getProperty("pos.pic.line2", "80, 250, 150, 200;2;") + commuPicList.get(0);
        String line5 = p.getProperty("pos.pic.line3", "350, 250, 150, 200;2;") + picList[i];
        bw.write(line3);
        bw.newLine();
        bw.write(line4);
        bw.newLine();
        bw.write(line5);
        bw.newLine();

      }
    }
    String pendstr = "###";
    bw.write(pendstr);
    bw.newLine();
    bw.close();
    fos.close();
  }

  //生成属性文件
  //@Test
  public void createConfig47(Properties p, String rootPath, String prjName, int model)
      throws IOException, NoSuchFieldException {

    /**
     * 客户名称：新励成  监测模板
     发布城市：北京
     发布内容：新励成
     发布位置：方舟苑/朝阳区北四环东路甲9号
     发布数量：9
     发布时间：2016.11.19-11.25          报告完成：2016.11.22

     */

    File file = new File(
        Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString());
    //if(!file.exists()) {
    //            file.createNewFile();
    //        }
    //String regionName = "金都杭城商务楼";
    FileOutputStream fos = new FileOutputStream(file);

    //        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
    BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));
    String f = Paths.get(rootPath, prjName + ".xlsx").toString();

    // String root = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
    File fo = new File(f);
    String prjname = getFileNameNoEx(fo.getName());
    String pstr = "#";
    bw.write(pstr);
    bw.newLine();
    List<ProjectDetail> result = getCommusByPrj(f);
    for (ProjectDetail pd : result) {
      String[] picList = getPicArray(pd.getName());
      List<String> commuPicList = getCommuPicList(pd.getName());
      int picCount = picList.length;

      int pageSize = (int) Math.ceil((picList.length + 1) / 3.0);
      for (int i = 0; i < pageSize; i++) {
        String line0 = "##";
        bw.write(line0);
        bw.newLine();
        String line3 = "";
        String line00="";
        if (model == 7) {

          String title = p.getProperty("content.text.line1",prjName);
          line3 =
              p.getProperty("pos.text.line1", "85, 50, 600, 200;1;") +
                  "客户名称：" + title + p.getProperty("font.line1") +
                  "###" + "发布城市：" + pd.getRegion() + p.getProperty("font.line1") +
                  "###" + "发布内容：" + prjName + p.getProperty("font.line1") +
                  "###" + "发布位置：" + pd.getName() + p.getProperty("font.line1");
        }else if(model==3){
          line00= p.getProperty("pos.text.line0")+p.getProperty("ppt.title.line0")+p.getProperty("font.title.line0");
          String title = p.getProperty("content.text.line1",prjName);
          line3 =
              p.getProperty("pos.text.line1", "85, 50, 600, 200;1;") +
                          "客户/品牌  ：" + title + p.getProperty("font.line1") +
                  "###" + "城市          ：" + pd.getRegion() + p.getProperty("font.line1") +
                  "###" + "媒体形式   ：框架" + p.getProperty("font.line1") +
                  "###" + "媒体点位   ：" + pd.getName() +p.getProperty("font.line1") +
                  "###" + "媒体数量   ：" + pd.getContractCount() + p.getProperty("font.line1") +
                  "###" + "发 布 期    ：" + pd.getDate1() + "                   上刊日期：" + pd.getDate2()
                  + p.getProperty("font.line1");
        }
        else {

          line3 =
              p.getProperty("pos.text.line1", "85, 50, 600, 200;1;") + "客户名称：" + prjName + p.getProperty("font.line1") +
                  "###" + "发布城市：" + pd.getRegion() + p.getProperty("font.line1") +
                  "###" + "发布内容：" + prjName + p.getProperty("font.line1") +
                  "###" + "发布位置：" + pd.getName() + "/" + pd.getPositionDesc() + p.getProperty("font.line1") +
                  "###" + "发布数量：" + pd.getContractCount() + p.getProperty("font.line1") +
                  "###" + "发布时间：" + pd.getDate1() + "              报告完成  ：" + pd.getDate2()
                  + p.getProperty("font.line1");

        }
        String lineBuilding = "";
        String line4 = "";
        String line5 = "";
        if (picCount % 3 == 0) {
          if (i == 0) {
            lineBuilding =
                p.getProperty("pos.pic.line2", "100,250,150,200;2;") + commuPicList.get(0);
            //line4 = "275,300,150,200;2;" + picList[i * 3];
            line5 = p.getProperty("pos.pic.line4", "450,250,150,200;2;") + picList[i * 3];
          } else if (i == 1) {
            lineBuilding =
                p.getProperty("pos.pic.line2", "100,250,150,200;2;") + picList[i * 3 - 2];
            //line4 = "275,250,150,200;2;" + picList[i * 2];
            line5 = p.getProperty("pos.pic.line4", "450,250,150,200;2;") + picList[i * 3 - 1];
          } else {
            lineBuilding =
                p.getProperty("pos.pic.line2", "100,250,150,200;2;") + picList[(i - 1) * 3];
            line4 = p.getProperty("pos.pic.line3", "275,250,150,200;2;") + picList[(i - 1) * 3 + 1];
            if (((i - 1) * 3 + 2) <= picCount - 1) {
              line5 =
                  p.getProperty("pos.pic.line4", "450,250,150,200;2;") + picList[(i - 1) * 3 + 2];
            }
          }
        } else if (picCount % 3 == 1) {

          if (i == 0) {
            lineBuilding =
                p.getProperty("pos.pic.line2", "100,250,150,200;2;") + commuPicList.get(0);
            //line4 = "275,300,150,200;2;" + picList[i * 3];
            line5 = p.getProperty("pos.pic.line4", "450,250,150,200;2;") + picList[i * 3];
          } else {
            lineBuilding =
                p.getProperty("pos.pic.line2", "100,250,150,200;2;") + picList[i * 3 - 2];
            line4 = p.getProperty("pos.pic.line3", "275,250,150,200;2;") + picList[i * 3 - 1];
            line5 = p.getProperty("pos.pic.line4", "450,250,150,200;2;") + picList[i * 3];
          }
        } else if (picCount % 3 == 2) {

          if (i == 0) {
            lineBuilding =
                p.getProperty("pos.pic.line2", "100,250,150,200;2;") + commuPicList.get(0);
            line4 = p.getProperty("pos.pic.line3", "275,250,150,200;2;") + picList[i * 3];
            line5 = p.getProperty("pos.pic.line4", "450,250,150,200;2;") + picList[i * 3 + 1];
          } else {
            lineBuilding =
                p.getProperty("pos.pic.line2", "100,250,150,200;2;") + picList[i * 3 - 1];
            line4 = p.getProperty("pos.pic.line3", "275,250,150,200;2;") + picList[i * 3];
            line5 = p.getProperty("pos.pic.line4", "450,250,150,200;2;") + picList[i * 3 + 1];
          }

        } else {

        }
        //line4 = "80, 300, 150, 150;2;" + picList[i * 2];
        //line5 = "350, 300, 150, 150;2;" + picList[i * 2 + 1];
        if(line00!=null&& !"".equals(line00)){
          bw.write(line00);
          bw.newLine();
        }
        if (line3 != null && !"".equals(line3)) {
          bw.write(line3);
          bw.newLine();
        }
        if (lineBuilding != null && !"".equals(lineBuilding)) {
          bw.write(lineBuilding);
          bw.newLine();
        }
        if (line4 != null && !"".equals(line4)) {
          bw.write(line4);
          bw.newLine();
        }
        if (line5 != null && !"".equals(line5)) {
          bw.write(line5);
          bw.newLine();
        }
      }
    }
    String pendstr = "###";
    bw.write(pendstr);
    bw.newLine();
    bw.close();
    fos.close();
  }

  /*
  社区位置：朝阳区东八里庄甘露园南里
合同规定：5
实际发布：5

   */
  //@Test
  public void createConfig6(Properties p, String rootPath, String prjName)
      throws IOException, NoSuchFieldException {
    //吉利

    File file = new File(
        Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString());
    //if(!file.exists()) {
    //            file.createNewFile();
    //        }
    //String regionName = "金都杭城商务楼";
    FileOutputStream fos = new FileOutputStream(file);

    //        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
    BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));

    String f = Paths.get(rootPath, prjName + ".xlsx").toString();

    //String root = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
    File fo = new File(f);
    String prjname = getFileNameNoEx(fo.getName());
    String pstr = "#";
    bw.write(pstr);
    bw.newLine();
    List<ProjectDetail> result = getCommusByPrj(f);
    ProjectDetail[] resulta = getPrjArray(result);
    int pageSize = (int) (Math.ceil(resulta.length / 2.0));
    for (int m = 0; m < pageSize; m++) {
      ProjectDetail pd1 = resulta[2*m];
      ProjectDetail pd2 = null;
      if ((2*m + 1) <= (resulta.length - 1)) {
        pd2 = resulta[2*m + 1];
      }

      String[] picList1 = getPicArray(pd1.getName());
      String[] picList2 = getPicArray(pd2.getName());

      List<String> commuPicList1 = getCommuPicList(pd1.getName());
      List<String> commuPicList2 = getCommuPicList(pd2.getName());
      //            int pageSize = picList.length / 2;

      //            for (int i = 0; i < pageSize; i++) {
      String line0 = "##";
      bw.write(line0);
      bw.newLine();
      String line01 = p.getProperty("pos.pic.line01", "30, 30, 150, 200;2;") + Paths.get(rootPath,"ASHARE","bar641.png").toString();
      String line1 = p.getProperty("pos.pic.line1", "30, 30, 150, 200;2;") + commuPicList1.get(0);
      String line2 = p.getProperty("pos.text.line2", "185, 30, 155, 30;1;") + pd1.getName()+p.getProperty("font.line2");

      String line631=p.getProperty("pos.text.line31")+ "社区位置：" + pd1.getPositionDesc()+p.getProperty("font.line31");
      String line632=p.getProperty("pos.text.line32")+ "合同规定：" +  pd1.getContractCount()+p.getProperty("font.line32");
      String line633=p.getProperty("pos.text.line33")+ "实际发布：" +pd1.getRealCount()+p.getProperty("font.line33");
      /**
      String line3 =
          p.getProperty("pos.text.line3", "185, 60, 155, 200;1;") + "社区位置：" + pd1.getPositionDesc()+p.getProperty("font.line3")
              +
              "###" + "合同规定：" + pd1.getContractCount()+p.getProperty("font.line3") +
              "###" + "实际发布：" + pd1.getRealCount()+p.getProperty("font.line3");
     */
      String line4 = p.getProperty("pos.text.line4", "30, 300, 200, 30;1;") + "发布实景图： "+p.getProperty("font.line4");
      String line05 = p.getProperty("pos.pic.line05", "30, 30, 150, 200;2;")  + Paths.get(rootPath,"ASHARE","bar643.png").toString();
      String line5 = p.getProperty("pos.pic.line5", "200, 340, 150, 200;2;") + picList1[0];
      String line011 = p.getProperty("pos.pic.line06", "421.00, 93.50, 0.00, 441.88;2;")  + Paths.get(rootPath,"ASHARE","bar642.png").toString();
      String line11 = p.getProperty("pos.pic.line6", "421.00, 93.50, 0.00, 441.88;2;") + commuPicList2.get(0);

      String line12 = p.getProperty("pos.pic.line12", "421.00, 93.50, 0.00, 441.88;2;") + Paths
          .get(rootPath,  "ASHARE", "bar6.png").toString();

      String line22 = p.getProperty("pos.text.line7", "610, 30, 155, 30;1;") + pd2.getName()+p.getProperty("font.line7");
      /**
      String line33 =
          p.getProperty("pos.text.line8", "610, 60, 155, 200;1;") + "社区位置：" + pd2.getPositionDesc()+p.getProperty("font.line8")
              +
              "###" + "合同规定：" + pd2.getContractCount()+p.getProperty("font.line8") +
              "###" + "实际发布：" + pd2.getRealCount()+p.getProperty("font.line8");
  */
          String line831=p.getProperty("pos.text.line81")+ "社区位置：" + pd2.getPositionDesc()+p.getProperty("font.line81");
      String line832=p.getProperty("pos.text.line82")+ "合同规定：" +  pd2.getContractCount()+p.getProperty("font.line82");
      String line833=p.getProperty("pos.text.line83")+ "实际发布：" +pd2.getRealCount()+p.getProperty("font.line83");
      String line44 = p.getProperty("pos.text.line9", "460, 300, 200, 30;1;") + "发布实景图： "+p.getProperty("font.line9");
      String line055 = p.getProperty("pos.pic.line010", "600, 340, 150, 200;2;") + Paths.get(rootPath,"ASHARE","bar644.png").toString();
      String line55 = p.getProperty("pos.pic.line10", "600, 340, 150, 200;2;") + picList2[0];

      bw.write(line01);
      bw.newLine();
      bw.write(line1);
      bw.newLine();
      bw.write(line2);
      bw.newLine();
      bw.write(line631);
      bw.newLine();
      bw.write(line632);
      bw.newLine();
      bw.write(line633);
      bw.newLine();
      bw.write(line831);
      bw.newLine();
      bw.write(line832);
      bw.newLine();
      bw.write(line833);
      bw.newLine();
      bw.write(line4);
      bw.newLine();
      bw.write(line05);
      bw.newLine();
      bw.write(line5);
      bw.newLine();
      bw.write(line12);
      bw.newLine();
      bw.write(line011);
      bw.newLine();
      bw.write(line11);
      bw.newLine();
      bw.write(line22);
      bw.newLine();
    //  bw.write(line33);
     // bw.newLine();
      bw.write(line44);
      bw.newLine();
      bw.write(line055);
      bw.newLine();
      bw.write(line55);
      bw.newLine();
      //       }
    }
    String pendstr = "###";
    bw.write(pendstr);
    bw.newLine();
    bw.close();
    fos.close();
  }

  /*
  发布品牌：广汽本田                                    发布城市：广州
发布内容：雅阁                                           发布位置：XXX小区/海珠区江南大道南384号
发布数量：XXXX                                         发布时间： 2016年10月08日-2016年10月14日

   */
  // @Test
  public void createConfig8(Properties p, String rootPath, String prjName)
      throws IOException, NoSuchFieldException {
    //上版
    File file = new File(
        Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString());
    //if(!file.exists()) {
    //            file.createNewFile();
    //        }
    //String regionName = "金都杭城商务楼";
    FileOutputStream fos = new FileOutputStream(file);

    //        BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos));
    BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));
    String f = rootPath + prjName + ".xlsx";

    //String root = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
    File fo = new File(f);
    String prjname = getFileNameNoEx(fo.getName());
    String pstr = "#";
    bw.write(pstr);
    bw.newLine();
    List<ProjectDetail> result = getCommusByPrj(f);
    for (ProjectDetail pd : result) {
      String[] picList = getPicArray(pd.getName());
      List<String> commuPicList = getCommuPicList(pd.getName());
      //int pageSize = picList.length / 2;
      int picCount = picList.length;
      int pageSize = (int) (Math.ceil((picList.length+1) / 3.0));
      for (int i = 0; i < pageSize; i++) {
        String line0 = "##";
        bw.write(line0);
        bw.newLine();
        //String line1 = "30, 30, 250, 250;2;" + commuPicList.get(0);
        // String line2 = "30, 300, 200, 30;1;" + pd.getName();
        String line1 = p.getProperty("pos.text.line1", "70,5,150,30;1;") + "上刊报告"+p.getProperty("font.line1");

        String line2 = p.getProperty("pos.pic.line2", "0,52,800,5;2;") + Paths
            .get(rootPath,  "ASHARE", "bar8.png").toString();
        String brand = p.getProperty("content.text.line35",prjName);
        String line31 =  p.getProperty("pos.text.line31", "50, 50, 500, 200;1;") + "发布品牌：" + prjName +p.getProperty("font.line31");
        String line32 =  p.getProperty("pos.text.line32", "50, 50, 500, 200;1;") + "发布城市："  +p.getProperty("font.line32");
        String line33 =  p.getProperty("pos.text.line33", "50, 50, 500, 200;1;") +  pd.getRegion()  +p.getProperty("font.line33");
        String line34 =  p.getProperty("pos.text.line34", "50, 50, 500, 200;1;") + "发布内容：" + brand +p.getProperty("font.line34");
        String line35 =  p.getProperty("pos.text.line35", "50, 50, 500, 200;1;") + "发布位置：" +p.getProperty("font.line35");
        String line36 =  p.getProperty("pos.text.line36", "50, 50, 500, 200;1;") + pd.getName() + "/" + pd.getPositionDesc()  +p.getProperty("font.line36");
        String line37 =  p.getProperty("pos.text.line37", "50, 50, 500, 200;1;") + "发布数量："+p.getProperty("font.line37");
        String line38 =  p.getProperty("pos.text.line38", "50, 50, 500, 200;1;") +  pd.getContractCount() +p.getProperty("font.line38");
        String line39 =  p.getProperty("pos.text.line39", "50, 50, 500, 200;1;") + "发布时间：" + pd.getDate1() +p.getProperty("font.line39");
      /**
        String line3 = p.getProperty("pos.text.line3", "50, 50, 500, 200;1;") + "发布品牌：" + prjName
            + "          发布城市：" + pd.getRegion()+// +p.getProperty("font.line3")+
              "###" + "发布内容：" + prjName + "          发布位置：" + pd.getName() + "/" + pd
            .getPositionDesc()  +//p.getProperty("font.line3")+
            "###" + "发布数量：" + pd.getContractCount() + "                    发布时间：" + pd.getDate1() +p.getProperty("font.line3");
       */



      //////////////////////////////////////////////////////////////////
      /**
        String line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") + commuPicList.get(0);
        String line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") + picList[2*i];
        String line6 ="";
        if((2*i+1)<=picList.length-1) {
          line6 = p.getProperty("pos.pic.line6", "450,250,150,200;2;") + picList[2*i + 1];

        }else{
          line5="";
          line6 = p.getProperty("pos.pic.line6", "275,250,150,200;2;") + picList[2*i];


        }
       */
////////////////////////////////////////////
        String line4="";
        String line5="";
        String line6="";

        if (picCount % 3 == 0) {
          if (i == 0) {
            line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") +  commuPicList.get(0);
        //    line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") +  picList[i * 3 - 1];
            line6 = p.getProperty("pos.pic.line6", "275,250,150,200;2;") + picList[i * 3];


          } else if (i == 1) {
            line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") +  picList[i * 3 - 2];
         //   line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") + picList[2*i];
            line6 =p.getProperty("pos.pic.line5", "275,250,150,200;2;") +  picList[i * 3 - 1];

          } else {
            line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") +  picList[(i - 1) * 3];
            line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") +  picList[(i - 1) * 3 + 1];
            if (((i - 1) * 3 + 2) <= picCount - 1) {
              line6 = p.getProperty("pos.pic.line6", "275,250,150,200;2;") +  picList[(i - 1) * 3 + 2];
            }

          }
        } else if (picCount % 3 == 1) {

          if (i == 0) {
            line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") + commuPicList.get(0);
        //    line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") +  picList[(i - 1) * 3 + 1];
            line6 = p.getProperty("pos.pic.line6", "275,250,150,200;2;") +  picList[i * 3];

          } else {
            line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") + picList[i * 3 - 2];
                line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") +  picList[i * 3 - 1];
            line6 = p.getProperty("pos.pic.line6", "275,250,150,200;2;") + picList[i * 3];

          }
        } else if (picCount % 3 == 2) {

          if (i == 0) {
            line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") + commuPicList.get(0);
            line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") + picList[i * 3];
            line6 = p.getProperty("pos.pic.line6", "275,250,150,200;2;") + picList[i * 3 + 1];

          } else {

            line4 = p.getProperty("pos.pic.line4", "100,250,150,200;2;") +  picList[i * 3 - 1];
            line5 = p.getProperty("pos.pic.line5", "275,250,150,200;2;") + picList[i * 3];
            line6 = p.getProperty("pos.pic.line6", "275,250,150,200;2;") +picList[i * 3 + 1];
          }

        } else {

        }

        //   String line4 = "350, 230, 150, 150;2;" + picList[i * 2];
        //                String line5 = "510, 230, 150, 150;2;" + picList[i * 2 + 1];
        bw.write(line1);
        bw.newLine();
        bw.write(line2);
        bw.newLine();
      //  bw.write(line3);
        bw.write(line31);
        bw.newLine();
        bw.write(line32);
        bw.newLine();
        bw.write(line33);
        bw.newLine();
        bw.write(line34);
        bw.newLine();
        bw.write(line35);
        bw.newLine();
        bw.write(line36);
        bw.newLine();
        bw.write(line37);
        bw.newLine();
        bw.write(line38);
        bw.newLine();
        bw.write(line39);
        bw.newLine();
        if(line4!=null&&!"".equals(line4))
        bw.write(line4);
        bw.newLine();
        if(!"".equals(line5)) {
          bw.write(line5);
          bw.newLine();
        }
        if(!"".equals(line6)) {
          bw.write(line6);
          bw.newLine();
        }
      }
    }
    String pendstr = "###";
    bw.write(pendstr);
    bw.newLine();
    bw.close();
    fos.close();
  }

  //生成属性文件
  //@Test
  public void createConfig(Properties p, String rootPath, String prjName)
      throws IOException, NoSuchFieldException {
    //KFC-北京

    File file = new File(
        Paths.get(rootPath.toString(), prjName, "conf", "ppcloud2.properties").toString());
    //if(!file.exists()) {
    //            file.createNewFile();
    //        }
    //String regionName = "金都杭城商务楼";
    FileOutputStream fos = new FileOutputStream(file);
    BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));

    String f = Paths.get(rootPath, prjName + ".xlsx").toString();

    //String root = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects/";
    File fo = new File(f);
    String prjname = getFileNameNoEx(fo.getName());
    String pstr = "#";
    bw.write(pstr);
    bw.newLine();
    List<ProjectDetail> result = getCommusByPrj(f);
    for (ProjectDetail pd : result) {
      String[] picList = getPicArray(pd.getName());
      List<String> commuPicList = getCommuPicList(pd.getName());
      int pageSize = picList.length / 2;
      for (int i = 0; i < pageSize; i++) {
        String line0 = "##";
        bw.write(line0);
        bw.newLine();
        //String line01= p.getProperty("pos.pic.line0")+Paths.get(rootPath,prjName,"conf","bar.png").toString();
        String line01= p.getProperty("pos.pic.line0")+Paths.get(rootPath,"ASHARE","bar1.png").toString();
        String line1 = p.getProperty("pos.pic.line1", "30, 30, 250, 250;2;") + commuPicList.get(0);
        String line2 = p.getProperty("pos.text.line2", "30, 300, 200, 30;1;") + pd.getName() + p
            .getProperty("font.line2", "16:宋体:255,0,0:1:0");
        String line3 =
            p.getProperty("pos.text.line3", "350, 20, 300, 200;1;") + "社区位置：" + pd.getPositionDesc()
                + p.getProperty("font.line3", "16:宋体:255,0,0:1:0") +
                "###" + "社区属性及人口：" + pd.getCommunityClassify()+", "+pd.getCommunityScale() + p
                .getProperty("font.line3", "16:宋体:255,0,0:1:0") +
                "###" + "用户描述：" + pd.getAudiences() + p
                .getProperty("font.line3", "16:宋体:255,0,0:1:0") +
                "###" + "入住率：" + pd.getOccupyRate() + p
                .getProperty("font.line3", "16:宋体:255,0,0:1:0") +
                "###" + "楼层：" + pd.getStories() + p.getProperty("font.line3", "16:宋体:255,0,0:1:0") +
                "###" + "合同规定：" + pd.getContractCount() + p
                .getProperty("font.line3", "16:宋体:255,0,0:1:0") +
                "###" + "实际发布：" + pd.getRealCount() + p
                .getProperty("font.line3", "16:宋体:255,0,0:1:0");
        String line4 = p.getProperty("pos.pic.line4", "350, 230, 150, 150;2;") + picList[i * 2];
        String line5 = p.getProperty("pos.pic.line5", "510, 230, 150, 150;2;") + picList[i * 2 + 1];
        bw.write(line01);
        bw.newLine();
        bw.write(line1);
        bw.newLine();
        bw.write(line2);
        bw.newLine();
        bw.write(line3);
        bw.newLine();
        bw.write(line4);
        bw.newLine();
        bw.write(line5);
        bw.newLine();
      }
    }
    String pendstr = "###";
    bw.write(pendstr);
    bw.newLine();
    bw.close();
    fos.close();
  }

  //@Test
  public void operImg2011() throws IOException {

    TextShape.TextDirection tds[] = {
        TextShape.TextDirection.HORIZONTAL,
        TextShape.TextDirection.VERTICAL,
        TextShape.TextDirection.VERTICAL_270,
        // TextDirection.STACKED is not supported on HSLF
    };

    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(templatePath.toString()));
    //7.6,7.6  7.4 5.6
    // extract all pictures contained in the presentation
    int idx = 1;

    for (XSLFPictureData pict : ppt.getPictureData()) {
      // picture data
      byte[] data = pict.getData();

      PictureData.PictureType type = pict.getType();
      String ext = type.extension;
      FileOutputStream out = new FileOutputStream("pict_" + idx + ext);
      out.write(data);
      out.close();
      idx++;
    }

    LoadProperties lp = new LoadProperties();
    for (SlideEntity slideEntity : lp.getPPTBean(p, rootPath.toString()).getSlideEntitiesList()) {
      createSlide(ppt, slideEntity);
    }
    /** simon2
     //table data
     String[][] data = {
     {"INPUT FILE", "NUMBER OF RECORDS"},
     {"Item File", "11,559"},
     {"Vendor File", "300"},
     {"Purchase History File", "10,000"},
     {"Total # of requisitions", "10,200,038"}
     };

     XSLFSlide slide = ppt.createSlide();
     //create a table of 5 rows and 2 columns

     TableShape<?, ?> tbl1 = slide.createTable(2, 2);
     tbl1.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));

     for (int i = 0; i < tbl1.getNumberOfRows(); i++) {
     for (int j = 0; j < tbl1.getNumberOfColumns(); j++) {
     TableCell<?, ?> cell = tbl1.getCell(i, j);
     if (i == 0) {
     cell.setBorderColor(TableCell.BorderEdge.left, ColorUIResource.CYAN);

     cell.setText("城市");
     } else {
     cell.setText(data[i][j]);
     }


     /////////////////////
     XSLFSlide slide2 = ppt.createSlide();
     XSLFTable tbl = slide2.createTable();
     tbl.setAnchor(new Rectangle2D.Double(50, 50, 450, 300));

     int numColumns = 3;
     int numRows = 5;
     XSLFTableRow headerRow = tbl.addRow();
     headerRow.setHeight(50);
     // header
     for (int i1 = 0; i1 < numColumns; i1++) {
     XSLFTableCell th = headerRow.addCell();
     XSLFTextParagraph p = th.addNewTextParagraph();
     p.setTextAlign(TextParagraph.TextAlign.CENTER);
     XSLFTextRun r = p.addNewTextRun();
     r.setText("Header " + (i1 + 1));
     r.setBold(true);
     r.setFontColor(Color.WHITE);
     th.setFillColor(Color.CYAN);
     th.setBorderWidth(TableCell.BorderEdge.bottom, 2);
     th.setBorderWidth(TableCell.BorderEdge.left, 2);
     th.setBorderWidth(TableCell.BorderEdge.top, 2);
     th.setBorderWidth(TableCell.BorderEdge.right, 2);

     th.setBorderColor(TableCell.BorderEdge.bottom, Color.cyan);
     th.setBorderColor(TableCell.BorderEdge.top, Color.cyan);
     th.setBorderColor(TableCell.BorderEdge.left, Color.cyan);
     th.setBorderColor(TableCell.BorderEdge.right, Color.cyan);
     tbl.setColumnWidth(i1, 150);  // all columns are equally sized
     }

     // rows

     for (int rownum = 0; rownum < numRows; rownum++) {
     XSLFTableRow tr = tbl.addRow();
     tr.setHeight(50);
     // header
     for (int i2 = 0; i2 < numColumns; i2++) {
     XSLFTableCell cell2 = tr.addCell();
     XSLFTextParagraph p = cell2.addNewTextParagraph();
     XSLFTextRun r = p.addNewTextRun();

     cell2.setBorderWidth(TableCell.BorderEdge.bottom, 2);
     cell2.setBorderWidth(TableCell.BorderEdge.top, 2);
     cell2.setBorderWidth(TableCell.BorderEdge.left, 2);
     cell2.setBorderWidth(TableCell.BorderEdge.right, 2);

     cell2.setBorderColor(TableCell.BorderEdge.bottom, Color.cyan);
     cell2.setBorderColor(TableCell.BorderEdge.top, Color.cyan);
     cell2.setBorderColor(TableCell.BorderEdge.right, Color.cyan);
     cell2.setBorderColor(TableCell.BorderEdge.left, Color.cyan);

     r.setText("Cell " + (i2 + 1));
     if (rownum % 2 == 0)
     cell2.setFillColor(Color.WHITE);
     else
     cell2.setFillColor(Color.YELLOW);

     }

     }

     }
     }
     */
    //set table borders
    /**
     Line border = tbl1.createBorder();
     border.setLineColor(Color.black);
     border.setLineWidth(1.0);
     table.setAllBorders(border);

     //set width of the 1st column
     table.setColumnWidth(0, 300);
     //set width of the 2nd column
     table.setColumnWidth(1, 150);

     slide.addShape(table);
     table.moveTo(100, 100);
     */

    /** simon
     // add a new picture to this slideshow and insert it in a new slide
     XSLFPictureData pd = ppt.addPicture(new File("a1.png"), PictureData.PictureType.PNG);

     // set image position in the slide

     XSLFSlide slide = ppt.createSlide();
     Dimension dd = ppt.getPageSize();
     System.out.println(dd.getHeight() + ":" + dd.getWidth());
     XSLFPictureShape shape = slide.createPicture(pd);
     Rectangle2D rect = new Rectangle(10, 10, (540 * 3 / 10), 540 * 3 / 10);



     shape.setAnchor(rect);
     simon
     */

    /**
     ///////////////////
     // add a new picture to this slideshow and insert it in a new slide
     // add a new picture to this slideshow and insert it in a new slide
     XSLFPictureData pd2 = ppt.addPicture(new File("a2.png"), PictureData.PictureType.PNG);

     // set image position in the slide

     XSLFPictureShape shape2 = slide.createPicture(pd2);
     shape2.setAnchor(new java.awt.Rectangle(115, 150, 100, 150));


     XSLFPictureData pd3 = ppt.addPicture(new File("a3.png"), PictureData.PictureType.PNG);

     // set image position in the slide

     XSLFPictureShape shape3 = slide.createPicture(pd3);
     shape2.setAnchor(new java.awt.Rectangle(220, 150, 100, 150));

     */

    /** simon
     // now retrieve pictures containes in the first slide and save them on disk
     idx = 1;
     slide = ppt.getSlides().get(0);
     for (XSLFShape sh : slide.getShapes()) {
     if (sh instanceof XSLFPictureShape) {
     XSLFPictureShape pict = (XSLFPictureShape) sh;
     XSLFPictureData pictData = pict.getPictureData();
     byte[] data = pictData.getData();
     PictureData.PictureType type = pictData.getType();
     FileOutputStream out = new FileOutputStream("slide0_" + idx + type.extension);
     out.write(data);
     out.close();
     idx++;
     }
     }

     */

    FileOutputStream out = new FileOutputStream(
        Paths.get(rootPath.toString(), prjName, "out", prjName + ".pptx").toString());
    ppt.write(out);
    out.close();

  }

  public List<ProjectDetail> getBrief(InputStream is) {

    List<ErrorMsgVO> errorList = new ArrayList<ErrorMsgVO>();
    List<MetaData> metaDatasBrief = new ArrayList<MetaData>();
    List<ProjectDetail> resultBrief = new ArrayList<ProjectDetail>();
    try {
      ExcelUtilExt<ProjectDetail> eeBrief = new ExcelUtilExt<ProjectDetail>(is, ProjectDetail.class,
          this.offset);
      //brief
      if (eeBrief.getTitleCount(this.offset) == this.getTitleCountBrief) {
        metaDatasBrief.add(new MetaData("code", "编号"));
      } else {
        errorList.add(new ErrorMsgVO("第一行", "列数不匹配", ""));
        //return errorList;
      }
      metaDatasBrief.add(new MetaData("name", "项目名称"));
      metaDatasBrief.add(new MetaData("contractCount", "合同数量"));
      metaDatasBrief.add(new MetaData("realCount", "实际数量"));

      eeBrief.setMetaDatas(metaDatasBrief);
      errorList = eeBrief.parseHead(this.offsetBrief, 0, metaDatasBrief, errorList);

      try {
        eeBrief.parseContent(this.offsetBrief + 1, 0, resultBrief);
      } catch (InstantiationException e) {
        e.printStackTrace();
      } catch (IllegalAccessException e) {
        e.printStackTrace();
      }
      System.out.println(resultBrief.size());

    } catch (IOException e) {
      e.printStackTrace();
    } catch (NoSuchFieldException e) {
      e.printStackTrace();
    }
    return resultBrief;

  }

  public List<ProjectDetail> getMergeList(List<ProjectDetail> brief, List<ProjectDetail> details) {
    List<ProjectDetail> mergeList = new ArrayList<ProjectDetail>();
    for (int i = 0; i < details.size(); i++) {
      ProjectDetail pd = details.get(i);
      for (int j = 0; j < brief.size(); j++) {
        ProjectDetail pdd = brief.get(j);
        if (pd.getName().equals(pdd.getName())) {
          pd.setContractCount(pdd.getContractCount());
          pd.setRealCount(pdd.getRealCount());
        }
        mergeList.add(pd);
      }
    }
    return mergeList;
  }

  // @Test
  public void getCommusByPrjTest() throws IOException, NoSuchFieldException {

    String file = Paths.get(rootPath.toString(), prjName + ".xlsx").toString();
    InputStream is = new FileInputStream(new File(file));
    List<ErrorMsgVO> errorList = new ArrayList<ErrorMsgVO>();
    List<MetaData> metaDatas = new ArrayList<MetaData>();
    List<ProjectDetail> result = new ArrayList<ProjectDetail>();
    ExcelUtilExt<ProjectDetail> ee = new ExcelUtilExt<ProjectDetail>(is, ProjectDetail.class,
        this.offset);

    //detail
    if (ee.getTitleCount(this.offset) == this.titleCount) {
      metaDatas.add(new MetaData("code", "编号"));
    } else {
      errorList.add(new ErrorMsgVO("第一行", "列数不匹配", ""));
      //return errorList;
    }
    // metaDatas.add(new MetaData("code", "编号"));
    metaDatas.add(new MetaData("name", "项目名称"));
    metaDatas.add(new MetaData("region", "区域"));
    metaDatas.add(new MetaData("positionDesc", "位 置 描 述 "));
    metaDatas.add(new MetaData("communityClassify", "社区分类"));
    metaDatas.add(new MetaData("avgPrice", "租售均价（人民币）"));
    metaDatas.add(new MetaData("communityScale", "社区居住规模（人）"));
    metaDatas.add(new MetaData("occupyRate", "入住率％"));
    metaDatas.add(new MetaData("audiences", "各社区内受众描述"));
    metaDatas.add(new MetaData("stories", "楼层"));
    metaDatas.add(new MetaData("unitCount", "门洞数"));
    metaDatas.add(new MetaData("liftCount", "电梯总数"));
    metaDatas.add(new MetaData("waitRoomCount", "等候厅数"));
    metaDatas.add(new MetaData("contractCount", "合同数量"));
    metaDatas.add(new MetaData("realCount", "实际数量"));
    metaDatas.add(new MetaData("buildingDetail", "楼号细分"));

    ee.setMetaDatas(metaDatas);
    int sheetNo = 0;
    if (ee.getSheetNum(true) >= 2) {
      sheetNo = 1;
    }
    errorList = ee.parseHead(this.offset, sheetNo, metaDatas, errorList);
    try {
      ee.parseContent(this.offset + 1, sheetNo, result);
    } catch (InstantiationException e) {
      e.printStackTrace();
    } catch (IllegalAccessException e) {
      e.printStackTrace();
    }
    System.out.println(result.size() + result.get(0).getName());
    List<ProjectDetail> ret = new ArrayList<ProjectDetail>();
    if (ee.getSheetNum(true) >= 2) {
      ret = getMergeList(getBrief(is), result);
    }
    if (ret != null && ret.size() > 0) {
      System.out.println(ret.size() + ":" + ret.get(0).getName());
    }

  }

  public List<ProjectDetail> getCommusByPrj(String file) throws IOException, NoSuchFieldException {

    InputStream is = new FileInputStream(new File(file));
    List<ErrorMsgVO> errorList = new ArrayList<ErrorMsgVO>();
    List<MetaData> metaDatas = new ArrayList<MetaData>();
    List<ProjectDetail> result = new ArrayList<ProjectDetail>();
    ExcelUtilExt<ProjectDetail> ee = new ExcelUtilExt<ProjectDetail>(is, ProjectDetail.class,
        this.offset);

    //detail
    if (ee.getTitleCount(this.offset) == this.titleCount) {
      metaDatas.add(new MetaData("code", "编号"));
    } else {
      errorList.add(new ErrorMsgVO("第一行", "列数不匹配", ""));
      //return errorList;
    }
    // metaDatas.add(new MetaData("code", "编号"));
    metaDatas.add(new MetaData("name", "项目名称"));
    metaDatas.add(new MetaData("region", "区域"));
    metaDatas.add(new MetaData("positionDesc", "位 置 描 述 "));
    metaDatas.add(new MetaData("communityClassify", "社区分类"));
    metaDatas.add(new MetaData("avgPrice", "租售均价（人民币）"));
    metaDatas.add(new MetaData("communityScale", "社区居住规模（人）"));
    metaDatas.add(new MetaData("occupyRate", "入住率％"));
    metaDatas.add(new MetaData("audiences", "各社区内受众描述"));
    metaDatas.add(new MetaData("stories", "楼层"));
    metaDatas.add(new MetaData("unitCount", "门洞数"));
    metaDatas.add(new MetaData("liftCount", "电梯总数"));
    metaDatas.add(new MetaData("waitRoomCount", "等候厅数"));
    metaDatas.add(new MetaData("contractCount", "合同数量"));
    metaDatas.add(new MetaData("realCount", "实际数量"));
    metaDatas.add(new MetaData("buildingDetail", "楼号细分"));
    if (titleCount == 18) {
      metaDatas.add(new MetaData("date1", "广告发布日期"));
      metaDatas.add(new MetaData("date2", "监测日期"));
    }

    ee.setMetaDatas(metaDatas);
    int sheetNo = 0;
    if (ee.getSheetNum(true) >= 2) {
      sheetNo = 1;
    }
    errorList = ee.parseHead(this.offset, sheetNo, metaDatas, errorList);
    try {
      ee.parseContent(this.offset + 1, sheetNo, result);
    } catch (InstantiationException e) {
      e.printStackTrace();
    } catch (IllegalAccessException e) {
      e.printStackTrace();
    }
    System.out.println(result.size());
    List<ProjectDetail> ret = new ArrayList<ProjectDetail>();
    if (ee.getSheetNum(true) >= 2) {
      ret = getMergeList(getBrief(is), result);
    }
    if (ret != null && ret.size() > 0) {
      System.out.println(ret.size());
      return ret;
    }
    return result;

  }

  //  @Test
  public void readExcel() throws IOException, NoSuchFieldException {
    File f = excelPath.toFile();
    InputStream is = new FileInputStream(f);
    List<ErrorMsgVO> errorList = new ArrayList<ErrorMsgVO>();
    List<MetaData> metaDatas = new ArrayList<MetaData>();

    ExcelUtilExt<ProjectDetail> ee = new ExcelUtilExt<ProjectDetail>(is, ProjectDetail.class, 3);

    //detail
    if (ee.getTitleCount(this.offset) == this.titleCount) {
      metaDatas.add(new MetaData("code", "编号"));
    } else {
      errorList.add(new ErrorMsgVO("第一行", "列数不匹配", ""));
      //return errorList;
    }
    // metaDatas.add(new MetaData("code", "编号"));
    metaDatas.add(new MetaData("name", "项目名称"));
    metaDatas.add(new MetaData("region", "区域"));
    metaDatas.add(new MetaData("positionDesc", "位 置 描 述 "));
    metaDatas.add(new MetaData("communityClassify", "社区分类"));
    metaDatas.add(new MetaData("avgPrice", "租售均价（人民币）"));
    metaDatas.add(new MetaData("communityScale", "社区居住规模（人）"));
    metaDatas.add(new MetaData("occupyRate", "入住率％"));
    metaDatas.add(new MetaData("audiences", "各社区内受众描述"));
    metaDatas.add(new MetaData("stories", "楼层"));
    metaDatas.add(new MetaData("unitCount", "门洞数"));
    metaDatas.add(new MetaData("liftCount", "电梯总数"));
    metaDatas.add(new MetaData("waitRoomCount", "等候厅数"));
    metaDatas.add(new MetaData("contractCount", "合同数量"));
    metaDatas.add(new MetaData("realCount", "实际数量"));
    metaDatas.add(new MetaData("buildingDetail", "楼号细分"));
    List<ProjectDetail> result = new ArrayList<ProjectDetail>();
    ee.setMetaDatas(metaDatas);
    errorList = ee.parseHead(this.offset, 1, metaDatas, errorList);
    try {
      ee.parseContent(this.offset + 1, 1, result);
    } catch (InstantiationException e) {
      e.printStackTrace();
    } catch (IllegalAccessException e) {
      e.printStackTrace();
    }
    System.out.println(result.size());
    List<ProjectDetail> ret = new ArrayList<ProjectDetail>();
    if (ee.getSheetNum(true) >= 2) {
      ret = getMergeList(getBrief(is), result);
    }
    if (ret != null && ret.size() > 0) {
      System.out.println(ret.size());
    }

  }

  //@Test
  public void createNewSlide() {
    //create a new empty slide show
    XMLSlideShow ppt = new XMLSlideShow();

    //add first slide
    XSLFSlide blankSlide = ppt.createSlide();
    writeOut(ppt, "simon.pptx");
  }

  //@Test
  public void readAndAppend() throws IOException {
    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("simon.pptx"));

    //append a new slide to the end
    XSLFSlide blankSlide = ppt.createSlide();
    writeOut(ppt, "simon2.pptx");

  }

  //@Test
  public void operImg() throws IOException {

    //金都杭城商务楼
    /**
     社区位置：朝阳区CBD商圈高档公寓
     社区属性及人口：B-ap, 8000
     用户描述：①+②+③+④+⑤+⑦
     入住率：100%
     楼层：17-26
     合同规定：12
     实际发布：12

     */
    HSLFSlideShow ppt = new HSLFSlideShow(new HSLFSlideShowImpl("master.ppt"));

    // extract all pictures contained in the presentation
    int idx = 1;
    for (HSLFPictureData pict : ppt.getPictureData()) {
      // picture data
      byte[] data = pict.getData();

      PictureData.PictureType type = pict.getType();
      String ext = type.extension;
      FileOutputStream out = new FileOutputStream("pict_" + idx + ext);
      out.write(data);
      out.close();
      idx++;
    }

    // add a new picture to this slideshow and insert it in a new slide
    HSLFPictureData pd = ppt.addPicture(new File("a1.png"), PictureData.PictureType.PNG);

    HSLFPictureShape pictNew = new HSLFPictureShape(pd);

    // set image position in the slide
    pictNew.setAnchor(new Rectangle(10, 10, 100, 150));

    HSLFSlide slide = ppt.createSlide();
    slide.addShape(pictNew);

    ///////////////////
    // add a new picture to this slideshow and insert it in a new slide
    HSLFPictureData pd2 = ppt.addPicture(new File("a2.png"), PictureData.PictureType.PNG);

    HSLFPictureShape pictNew2 = new HSLFPictureShape(pd2);

    // set image position in the slide
    pictNew2.setAnchor(new Rectangle(115, 150, 100, 150));

    slide.addShape(pictNew2);
    ////////////////////
    HSLFPictureData pd3 = ppt.addPicture(new File("a3.png"), PictureData.PictureType.PNG);

    HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);

    // set image position in the slide
    pictNew3.setAnchor(new Rectangle(220, 150, 100, 150));

    slide.addShape(pictNew3);

    // now retrieve pictures containes in the first slide and save them on disk
    idx = 1;
    slide = ppt.getSlides().get(0);
    for (HSLFShape sh : slide.getShapes()) {
      if (sh instanceof HSLFPictureShape) {
        HSLFPictureShape pict = (HSLFPictureShape) sh;
        HSLFPictureData pictData = pict.getPictureData();
        byte[] data = pictData.getData();
        PictureData.PictureType type = pictData.getType();
        FileOutputStream out = new FileOutputStream("slide0_" + idx + type.extension);
        out.write(data);
        out.close();
        idx++;
      }
    }

    FileOutputStream out = new FileOutputStream("simon7.ppt");
    ppt.write(out);
    out.close();

  }

  public void createSlide(XMLSlideShow xmlSlideShow, SlideEntity slideEntity) throws IOException {

    List<ElementBean> elementBeanList = slideEntity.getElementBeanList();
    XSLFSlide slide = xmlSlideShow.createSlide();

    Dimension dim = xmlSlideShow.getPageSize();
    double pgheight = dim.getHeight();
    double pgwidth = dim.getWidth();
    for (ElementBean elementBean : elementBeanList) {
      if (elementBean.getType() == 1) {

        XSLFTextBox shape = slide.createTextBox();

        shape.setVerticalAlignment(VerticalAlignment.MIDDLE);

        String[] cont = elementBean.getContent().trim().split("###");
        if (cont != null && cont.length > 0) {
          for (String para : cont) {
            String[] paras = para.split(":");
            //ztbjys
            if(paras.length==9){
              String[] costr = paras[8].split(",");
              shape.setFillColor(new Color(Integer.parseInt(costr[0]),Integer.parseInt(costr[1]),Integer.parseInt(costr[2])));
            }
            XSLFTextParagraph p = shape.addNewTextParagraph();

            //hjj
            if(paras.length>=8) {
              System.out.println("set line spacing:"+paras[7]);

              p.setLineSpacing(Double.parseDouble(paras[7]));
            }
            if(paras!=null&& paras.length>=7){
              if(paras[6]=="1") {
                p.setTextAlign(TextParagraph.TextAlign.CENTER);
                p.setTextAlign(TextParagraph.TextAlign.CENTER);
              }

            }

            XSLFTextRun r1 = p.addNewTextRun();

            r1.setSubscript(false);
            r1.setText(paras[0]);

            if (paras != null && paras.length >= 6) {
              r1.setFontSize(Double.parseDouble(paras[1]));
             String[] fonts = paras[2].split("=");

              r1.setFontFamily(fonts[0]);
              ///simon
              CTTextCharacterProperties xxp=r1.getXmlObject().getRPr();
              CTTextFont cttf = xxp.getLatin();
              CTTextFont ctsym = xxp.getSym();
              CTTextFont eattf = xxp.getEa();
              CTTextFont ctcs = xxp.getCs();
              if(eattf==null){
                xxp.addNewEa();
//                xxp.setEa(cttf);
                 eattf = xxp.getEa();
              }
              if(ctcs==null){
                xxp.addNewCs();
                ctcs = xxp.getCs();
              }
              if(cttf==null){
                xxp.addNewLatin();
                cttf = xxp.getLatin();
              }
              if(ctsym==null){
               xxp.addNewSym();
               ctsym = xxp.getSym();
              }
              eattf.setTypeface(fonts[0]);
              //cttf.setTypeface(fonts[0]);
              if(fonts.length==2) {
                ctcs.setTypeface(fonts[1]);
                cttf.setTypeface(fonts[1]);
                ctsym.setTypeface(fonts[1]);
              }else{
                ctcs.setTypeface(fonts[0]);
                cttf.setTypeface(fonts[0]);
                ctsym.setTypeface(fonts[0]);
              }
              xxp.setEa(eattf);
              xxp.setLatin(cttf);
              xxp.setSym(ctsym);
              xxp.setCs(ctcs);
              ///

              String[] colorarr = paras[3].split(",");
              r1.setFontColor(
                  new Color(Integer.parseInt(colorarr[0]), Integer.parseInt(colorarr[1]),
                      Integer.parseInt(colorarr[2])));
              r1.setBold("1".equals(paras[4]));
              r1.setItalic("1".equals(paras[5]));

            } else {
              r1.setFontColor(Color.BLACK);
              r1.setFontFamily("黑体");
              r1.setFontSize(12.);

            }
          }
        }
        Rectangle2D rect = new Rectangle2D.Double(elementBean.getX(), elementBean.getY(),
            elementBean.getWidth(), elementBean.getHeight());
        /**
         {
         @Override public void setRect(double x, double y, double w, double h) {

         }

         @Override public int outcode(double x, double y) {
         return 0;
         }

         @Override public Rectangle2D createIntersection(Rectangle2D r) {
         return null;
         }

         @Override public Rectangle2D createUnion(Rectangle2D r) {
         return null;
         }

         @Override public double getX() {
         return 0;
         }

         @Override public double getY() {
         return 0;
         }

         @Override public double getWidth() {
         return 0;
         }

         @Override public double getHeight() {
         return 0;
         }

         @Override public boolean isEmpty() {
         return false;
         }
         };
         */
        shape.setAnchor(rect);

      } else if (elementBean.getType() == 2) {
        XSLFPictureData pd = xmlSlideShow
            .addPicture(new File(elementBean.getContent().trim()), PictureData.PictureType.PNG);

        Dimension dd = xmlSlideShow.getPageSize();
        System.out.println(dd.getHeight() + ":" + dd.getWidth());
        XSLFPictureShape shape = slide.createPicture(pd);
        Rectangle2D rect = new Rectangle2D.Double(elementBean.getX(), elementBean.getY(),
            elementBean.getWidth(), elementBean.getHeight());
        shape.setAnchor(rect);
      } else if (elementBean.getType() == 3) {

        XSLFTable tbl = slide.createTable();
        tbl.setAnchor(
            new Rectangle2D.Double(elementBean.getX(), elementBean.getY(), elementBean.getWidth(),
                elementBean.getHeight()));
        //tbl.setAnchor(new Rectangle2D.Double((pgwidth-elementBean.getWidth())/2, elementBean.getY(), elementBean.getWidth(), elementBean.getHeight()));
        System.out.println("pgwidth" + pgwidth + "tblwidth:" + elementBean.getWidth());
        System.out.println("pgheight:" + pgheight + "tblheight:" + elementBean.getHeight());

        XSLFTableRow headerRow = tbl.addRow();
        headerRow.setHeight(30);
        int[] width4 = {97,176,147,221};
        if (elementBean.getContent() != null) {
          String[] con = elementBean.getContent().split("!!!");
          if (null != con && con.length == 2) {
            String[] header = con[0].split("##");
            int numColumns = header.length;

            for (int i1 = 0; i1 < numColumns; i1++) {
              XSLFTableCell th = headerRow.addCell();
              th.setVerticalAlignment(VerticalAlignment.MIDDLE);
              XSLFTextParagraph p = th.addNewTextParagraph();
              p.setTextAlign(TextParagraph.TextAlign.CENTER);
              XSLFTextRun r = p.addNewTextRun();
              r.setText(header[i1]);
              r.setBold(true);
              //simon
              r.setFontFamily("微软雅黑");
              ///simon
              CTTextCharacterProperties xxp=r.getXmlObject().getRPr();
              CTTextFont cttf = xxp.getLatin();
              cttf.setTypeface("微软雅黑");
              xxp.setLatin(cttf);
              xxp.setEa(cttf);

              r.setFontSize(14.);
              r.setFontColor(Color.WHITE);
              th.setFillColor(new Color(0,176,240));
              th.setBorderWidth(TableCell.BorderEdge.bottom, 2);
              th.setBorderWidth(TableCell.BorderEdge.left, 2);
              th.setBorderWidth(TableCell.BorderEdge.top, 2);
              th.setBorderWidth(TableCell.BorderEdge.right, 2);

              th.setBorderColor(TableCell.BorderEdge.bottom, new Color(0,204,255));
              th.setBorderColor(TableCell.BorderEdge.top, new Color(0,204,255));
              th.setBorderColor(TableCell.BorderEdge.left, new Color(0,204,255));
              th.setBorderColor(TableCell.BorderEdge.right, new Color(0,204,255));

              tbl.setColumnWidth(i1, width4[i1]);  // all columns are equally sized

            }
            String[] rows = con[1].split("@@");
            for (String rowstr : rows) {

            }

            int numRows = rows.length;

            for (int rownum = 0; rownum < numRows; rownum++) {
              XSLFTableRow tr = tbl.addRow();

              tr.setHeight(50);
              // header
              String[] row = rows[rownum].split("##");
              for (int i2 = 0; i2 < numColumns; i2++) {
                XSLFTableCell cell2 = tr.addCell();
                cell2.setVerticalAlignment(VerticalAlignment.MIDDLE);
                XSLFTextParagraph p = cell2.addNewTextParagraph();
                p.setTextAlign(TextParagraph.TextAlign.CENTER);
                XSLFTextRun r = p.addNewTextRun();

                cell2.setBorderWidth(TableCell.BorderEdge.bottom, 2);
                cell2.setBorderWidth(TableCell.BorderEdge.top, 2);
                cell2.setBorderWidth(TableCell.BorderEdge.left, 2);
                cell2.setBorderWidth(TableCell.BorderEdge.right, 2);

                cell2.setBorderColor(TableCell.BorderEdge.bottom, new Color(0,204,255));
                cell2.setBorderColor(TableCell.BorderEdge.top, new Color(0,204,255));
                cell2.setBorderColor(TableCell.BorderEdge.right, new Color(0,204,255));
                cell2.setBorderColor(TableCell.BorderEdge.left, new Color(0,204,255));

                r.setText(row[i2]);
                //simon
                r.setFontFamily("微软雅黑");
                ///simon
                CTTextCharacterProperties xxp=r.getXmlObject().getRPr();
                CTTextFont cttf = xxp.getLatin();
                cttf.setTypeface("微软雅黑");
                xxp.setEa(cttf);
                xxp.setLatin(cttf);
                r.setFontSize(14.);

                if (rownum % 2 == 0)
                  cell2.setFillColor(Color.WHITE);
                else
                  cell2.setFillColor(Color.YELLOW);

              }

            }

          }
        }

      } else {

      }
    }

  }

  public String getProjectPath() {
    return "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects";
  }

  //@Test
  public void run() {
    String path = "/Users/wangqingwu/Projects/gen-pptx/pptgenerator/projects";

  }

  // @Test
  public void getProjectList() throws IOException {
    String prjdir = getProjectPath();
    Files.newDirectoryStream(
        Paths.get(prjdir),
        path -> Files.isDirectory(path))
        .forEach(System.out::println);

  }

  public void getCommunityOfProjectFromExcel(String excel) {

    System.out.println("come on here");
  }

  public static String getFileNameNoEx(String filename) {
    if ((filename != null) && (filename.length() > 0)) {
      int dot = filename.lastIndexOf('.');
      if ((dot > -1) && (dot < (filename.length()))) {
        return filename.substring(0, dot);
      }
    }
    return filename;
  }

  public void getCommunityOfProject(String prjPath,
      Consumer<? super Path> getCommunityOfProjectFromExcel) throws IOException {
    Files.newDirectoryStream(Paths.get(prjPath), path -> path.toString().endsWith(".xlsx"))
        .forEach(getCommunityOfProjectFromExcel);

  }

  //根据项目名称获取图片列表
  public List<String> getCommuPicList(String regionName) {

    List<String> ret = new ArrayList<String>();
    String pathPic = Paths.get(rootPath.toString(), prjName, regionName, "小区").toString();
    File f = new File(pathPic);
    if (f.exists() && f.isDirectory()) {
      String[] fileArray = f.list(new FilenameFilter() {
        @Override
        public boolean accept(File dir, String name) {
          System.out.println("commu:"+name);
          if (name.toLowerCase().endsWith("png") || name.toLowerCase().endsWith("jpg") || name
              .toLowerCase().endsWith("jpeg")) {
            return true;
          }
          return false;

        }
      });
      ret = Arrays.asList(fileArray);
    }

    ret = ret.stream().map(s -> getAbsPath(f, s)).collect(Collectors.toList());
    System.out.println(ret.size());
    return ret;

  }

  public String[] getPicArray(String regionName) {
    List<String> ret = getPic(regionName);
    final int size = ret.size();
    String[] arr = (String[]) ret.toArray(new String[size]);
    return arr;
  }

  public ProjectDetail[] getPrjArray(List<ProjectDetail> pd) {
    ProjectDetail[] pda = (ProjectDetail[]) pd.toArray(new ProjectDetail[pd.size()]);
    return pda;
  }

  //根据项目名称获取图片列表
  public List<String> getPic(String regionName) {

    List<String> ret = new ArrayList<String>();
    String pathPic = Paths.get(rootPath.toString(), prjName, regionName, "广告").toString();
    File f = new File(pathPic);
    if (f.exists() && f.isDirectory()) {
      String[] fileArray = f.list(new FilenameFilter() {
        @Override
        public boolean accept(File dir, String name) {
          if (name.toLowerCase().endsWith("png") || name.toLowerCase().endsWith("jpg") || name
              .toLowerCase().endsWith("jpeg")) {
            return true;
          }
          return false;

        }
      });
      ret = Arrays.asList(fileArray);
    }

    ret = ret.stream().map(s -> getAbsPath(f, s)).collect(Collectors.toList());
    System.out.println(ret.size());
    return ret;

  }

  public String getAbsPath(File f, String fname) {
    return Paths.get(f.getAbsolutePath(), fname).toString();

  }

  public List<String> getRegion() {
    String pathRegion = Paths.get(rootPath.toString(), prjName, "小区").toString();
    return null;
  }

  /**
   @Test public void operImg2011() throws IOException {

   TextShape.TextDirection tds[] = {
   TextShape.TextDirection.HORIZONTAL,
   TextShape.TextDirection.VERTICAL,
   TextShape.TextDirection.VERTICAL_270,
   // TextDirection.STACKED is not supported on HSLF
   };

   //金都杭城商务楼
   /**
   社区位置：朝阳区CBD商圈高档公寓
   社区属性及人口：B-ap, 8000
   用户描述：①+②+③+④+⑤+⑦
   入住率：100%
   楼层：17-26
   合同规定：12
   实际发布：12

   //

   XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(templatePath.toString()));
   //7.6,7.6  7.4 5.6
   // extract all pictures contained in the presentation
   int idx = 1;

   for (XSLFPictureData pict : ppt.getPictureData()) {
   // picture data
   byte[] data = pict.getData();

   PictureData.PictureType type = pict.getType();
   String ext = type.extension;
   FileOutputStream out = new FileOutputStream("pict_" + idx + ext);
   out.write(data);
   out.close();
   idx++;
   }

   LoadProperties lp = new LoadProperties();
   for (SlideEntity slideEntity : lp.getPPTBean().getSlideEntitiesList()) {
   createSlide(ppt, slideEntity);
   }

   //table data
   String[][] data = {
   {"INPUT FILE", "NUMBER OF RECORDS"},
   {"Item File", "11,559"},
   {"Vendor File", "300"},
   {"Purchase History File", "10,000"},
   {"Total # of requisitions", "10,200,038"}
   };

   XSLFSlide slide = ppt.createSlide();
   //create a table of 5 rows and 2 columns

   TableShape<?, ?> tbl1 = slide.createTable(2, 2);
   tbl1.setAnchor(new Rectangle2D.Double(50, 50, 200, 200));
   /**
   int col = 0;
   for (TextShape.TextDirection td : tds) {
   TableCell<?, ?> c = tbl1.getCell(0, col++);
   c.setTextDirection(td);
   c.setText("bla");
   }
   //
   for (int i = 0; i < tbl1.getNumberOfRows(); i++) {
   for (int j = 0; j < tbl1.getNumberOfColumns(); j++) {
   TableCell<?, ?> cell = tbl1.getCell(i, j);
   if (i == 0) {
   cell.setBorderColor(TableCell.BorderEdge.left, ColorUIResource.CYAN);
   cell.setText("城市");
   } else {
   cell.setText(data[i][j]);
   }


   /////////////////////
   XSLFSlide slide2 = ppt.createSlide();
   XSLFTable tbl = slide2.createTable();
   tbl.setAnchor(new Rectangle2D.Double(50, 50, 450, 300));

   int numColumns = 3;
   int numRows = 5;
   XSLFTableRow headerRow = tbl.addRow();
   headerRow.setHeight(50);
   // header
   for (int i1 = 0; i1 < numColumns; i1++) {
   XSLFTableCell th = headerRow.addCell();
   XSLFTextParagraph p = th.addNewTextParagraph();
   p.setTextAlign(TextParagraph.TextAlign.CENTER);
   XSLFTextRun r = p.addNewTextRun();
   r.setText("Header " + (i1 + 1));
   r.setBold(true);
   r.setFontColor(java.awt.Color.WHITE);
   th.setFillColor(java.awt.Color.CYAN);
   th.setBorderWidth(TableCell.BorderEdge.bottom, 2);
   th.setBorderWidth(TableCell.BorderEdge.left, 2);
   th.setBorderWidth(TableCell.BorderEdge.top, 2);
   th.setBorderWidth(TableCell.BorderEdge.right, 2);

   th.setBorderColor(TableCell.BorderEdge.bottom, java.awt.Color.cyan);
   th.setBorderColor(TableCell.BorderEdge.top, java.awt.Color.cyan);
   th.setBorderColor(TableCell.BorderEdge.left, java.awt.Color.cyan);
   th.setBorderColor(TableCell.BorderEdge.right, java.awt.Color.cyan);
   tbl.setColumnWidth(i1, 150);  // all columns are equally sized
   }

   // rows

   for (int rownum = 0; rownum < numRows; rownum++) {
   XSLFTableRow tr = tbl.addRow();
   tr.setHeight(50);
   // header
   for (int i2 = 0; i2 < numColumns; i2++) {
   XSLFTableCell cell2 = tr.addCell();
   XSLFTextParagraph p = cell2.addNewTextParagraph();
   XSLFTextRun r = p.addNewTextRun();

   cell2.setBorderWidth(TableCell.BorderEdge.bottom, 2);
   cell2.setBorderWidth(TableCell.BorderEdge.top, 2);
   cell2.setBorderWidth(TableCell.BorderEdge.left, 2);
   cell2.setBorderWidth(TableCell.BorderEdge.right, 2);

   cell2.setBorderColor(TableCell.BorderEdge.bottom, java.awt.Color.cyan);
   cell2.setBorderColor(TableCell.BorderEdge.top, java.awt.Color.cyan);
   cell2.setBorderColor(TableCell.BorderEdge.right, java.awt.Color.cyan);
   cell2.setBorderColor(TableCell.BorderEdge.left, java.awt.Color.cyan);

   r.setText("Cell " + (i2 + 1));
   if (rownum % 2 == 0)
   cell2.setFillColor(java.awt.Color.WHITE);
   else
   cell2.setFillColor(java.awt.Color.YELLOW);

   }

   }
   //////////////////////

   //XSLFTextRun<?> rt = cell.getTextParagraphs().get(0).getTextRuns()
   //rt.setFontFamily("Arial");
   //rt.setFontSize(10.);
   //cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
   //cell.setHorizontalCentered(true);
   }
   }

   //set table borders
   /**
   Line border = tbl1.createBorder();
   border.setLineColor(Color.black);
   border.setLineWidth(1.0);
   table.setAllBorders(border);

   //set width of the 1st column
   table.setColumnWidth(0, 300);
   //set width of the 2nd column
   table.setColumnWidth(1, 150);

   slide.addShape(table);
   table.moveTo(100, 100);
   */

  /** simon
   // add a new picture to this slideshow and insert it in a new slide
   XSLFPictureData pd = ppt.addPicture(new File("a1.png"), PictureData.PictureType.PNG);

   // set image position in the slide

   XSLFSlide slide = ppt.createSlide();
   Dimension dd = ppt.getPageSize();
   System.out.println(dd.getHeight() + ":" + dd.getWidth());
   XSLFPictureShape shape = slide.createPicture(pd);
   Rectangle2D rect = new Rectangle(10, 10, (540 * 3 / 10), 540 * 3 / 10);



   shape.setAnchor(rect);
   simon
   */

  /**
   ///////////////////
   // add a new picture to this slideshow and insert it in a new slide
   // add a new picture to this slideshow and insert it in a new slide
   XSLFPictureData pd2 = ppt.addPicture(new File("a2.png"), PictureData.PictureType.PNG);

   // set image position in the slide

   XSLFPictureShape shape2 = slide.createPicture(pd2);
   shape2.setAnchor(new java.awt.Rectangle(115, 150, 100, 150));


   XSLFPictureData pd3 = ppt.addPicture(new File("a3.png"), PictureData.PictureType.PNG);

   // set image position in the slide

   XSLFPictureShape shape3 = slide.createPicture(pd3);
   shape2.setAnchor(new java.awt.Rectangle(220, 150, 100, 150));

   */

  /**
   * simon
   * // now retrieve pictures containes in the first slide and save them on disk
   * idx = 1;
   * slide = ppt.getSlides().get(0);
   * for (XSLFShape sh : slide.getShapes()) {
   * if (sh instanceof XSLFPictureShape) {
   * XSLFPictureShape pict = (XSLFPictureShape) sh;
   * XSLFPictureData pictData = pict.getPictureData();
   * byte[] data = pictData.getData();
   * PictureData.PictureType type = pictData.getType();
   * FileOutputStream out = new FileOutputStream("slide0_" + idx + type.extension);
   * out.write(data);
   * out.close();
   * idx++;
   * }
   * }
   * <p>
   * //
   * <p>
   * FileOutputStream out = new FileOutputStream(Paths.get(rootPath.toString(), prjName, "out", prjName + ".pptx").toString());
   * ppt.write(out);
   * out.close();
   * <p>
   * }
   */
  //@Test
  public void createByLayout() throws IOException {
    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("layouts.pptx"));

    // first see what slide layouts are available :
    System.out.println("Available slide layouts:");
    for (XSLFSlideMaster master : ppt.getSlideMasters()) {
      for (XSLFSlideLayout layout : master.getSlideLayouts()) {
        System.out.println(layout.getType());
      }
    }

    XSLFSlideLayout detailedscorecard = null;
    for (XSLFSlideMaster master : ppt.getSlideMasters()) {
      for (XSLFSlideLayout layout1 : master.getSlideLayouts()) {
        System.out.println("0000" + layout1.getName());
        //if (layout1.getName().equals("Scorecard")) {
        //        detailedscorecard=layout1;
        //    }
      }
    }

    // blank slide
    XSLFSlide blankSlide = ppt.createSlide();

    // there can be multiple masters each referencing a number of layouts
    // for demonstration purposes we use the first (default) slide master
    XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);

    // title slide
    XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);
    // fill the placeholders
    XSLFSlide slide1 = ppt.createSlide(titleLayout);
    XSLFTextShape title1 = slide1.getPlaceholder(0);
    title1.setText("First Title");

    // title and content
    XSLFSlideLayout titleBodyLayout = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
    XSLFSlide slide2 = ppt.createSlide(titleBodyLayout);

    XSLFTextShape title2 = slide2.getPlaceholder(0);
    title2.setText("Second Title");

    XSLFTextShape body2 = slide2.getPlaceholder(1);
    body2.clearText(); // unset any existing text
    body2.addNewTextParagraph().addNewTextRun().setText("First paragraph");
    body2.addNewTextParagraph().addNewTextRun().setText("Second paragraph");
    body2.addNewTextParagraph().addNewTextRun().setText("Third paragraph");

    XSLFSlideLayout pic_tx = defaultMaster.getLayout(SlideLayout.PIC_TX);
    XSLFSlide slide3 = ppt.createSlide(pic_tx);

    XSLFTextShape pic = slide3.getPlaceholder(0);

    writeOut(ppt, "simon3.pptx");

  }

  /**
   * @Test public void testOOXML() throws InvalidFormatException, IOException {
   * XMLSlideShow pptx = new XMLSlideShow();
   * XSLFSlide slide = pptx.createSlide();
   * <p>
   * // you need to include ooxml-schemas:1.1 for this to work!!!
   * // otherwise an empty table will be created
   * // see https://issues.apache.org/bugzilla/show_bug.cgi?id=49934
   * XSLFTable table = slide.createTable();
   * table.setAnchor(new Rectangle2D.Double(50, 50, 500, 20));
   * <p>
   * XSLFTableRow row = table.addRow();
   * row.addCell().setText("Cell 1");
   * XSLFTableCell cell = row.addCell();
   * cell.setText("Cell 2");
   * <p>
   * <p>
   * CTBlipFillProperties blipPr = cell.getXmlObject().getTcPr().addNewBlipFill();
   * blipPr.setDpi(72);
   * // http://officeopenxml.com/drwPic-ImageData.php
   * CTBlip blib = blipPr.addNewBlip();
   * blipPr.addNewSrcRect();
   * CTRelativeRect fillRect = blipPr.addNewStretch().addNewFillRect();
   * fillRect.setL(30000);
   * fillRect.setR(30000);
   * <p>
   * PackagePartName partName = PackagingURIHelper.createPartName("/ppt/media/100px.gif");
   * PackagePart part = pptx.getPackage().createPart(partName, "image/gif");
   * OutputStream partOs = part.getOutputStream();
   * FileInputStream fis = new FileInputStream("src/test/resources/100px.gif");
   * byte buf[] = new byte[1024];
   * for (int readBytes; (readBytes = fis.read(buf)) != -1; partOs.write(buf, 0, readBytes));
   * fis.close();
   * partOs.close();
   * <p>
   * PackageRelationship prs = slide.getPackagePart().addRelationship(partName, TargetMode.INTERNAL, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
   * <p>
   * blib.setEmbed(prs.getId());
   * <p>
   * <p>
   * FileOutputStream fos = new FileOutputStream("test2.pptx");
   * pptx.write(fos);
   * fos.close();
   * }
   */

  public void writeOut(SlideShow ppt, String path) {
    FileOutputStream out = null;
    try {
      out = new FileOutputStream(path);
      ppt.write(out);
      out.close();
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    } finally {

    }
  }

  /**
   * 在一个幻灯片上话一个shape
   * <p>
   * 一个shape的位置同安卓里控件的位置
   */
  // @Test
  public void genSlide() {
    HSLFSlideShow ppt = new HSLFSlideShow();

    HSLFSlide slide = ppt.createSlide();

    FileOutputStream out = null;
    try {
      out = new FileOutputStream(outputPath.toString());
      ppt.write(out);
      out.close();
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    } finally {

    }

  }

  public void createSlide() {

  }

  // @Test
  public void addImage2Slide() throws IOException {
    XMLSlideShow ppt = new XMLSlideShow();
    XSLFSlide slide = ppt.createSlide();

    byte[] pictureData = IOUtils.toByteArray(new FileInputStream("1.jpeg"));

    XSLFPictureData pd = ppt.addPicture(pictureData, PictureData.PictureType.JPEG);
    XSLFPictureShape pic = slide.createPicture(pd);
    byte[] pictureData2 = IOUtils.toByteArray(new FileInputStream("2.jpeg"));

    XSLFPictureData pd2 = ppt.addPicture(pictureData2, PictureData.PictureType.JPEG);
    XSLFPictureShape pic2 = slide.createPicture(pd2);
    byte[] pictureData3 = IOUtils.toByteArray(new FileInputStream("3.png"));

    XSLFPictureData pd3 = ppt.addPicture(pictureData3, PictureData.PictureType.PNG);
    XSLFPictureShape pic3 = slide.createPicture(pd3);

    writeOut(ppt, "simon5.pptx");
  }

  // @Test
  public void readPicInSlide() throws IOException {
    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(tp1.toString()));
    // get slides
    for (XSLFSlide slide : ppt.getSlides()) {
      for (XSLFShape sh : slide.getShapes()) {
        // name of the shape
        String name = sh.getShapeName();
        System.out.println("-------" + name);
        // shapes's anchor which defines the position of this shape in the slide
        if (sh instanceof PlaceableShape) {
          Rectangle2D anchor = ((PlaceableShape) sh).getAnchor();
        }

        if (sh instanceof XSLFConnectorShape) {
          XSLFConnectorShape line = (XSLFConnectorShape) sh;
          // work with Line
        } else if (sh instanceof XSLFAutoShape) {
          XSLFAutoShape shape = (XSLFAutoShape) sh;
          for (XSLFTextParagraph xtp : shape.getTextParagraphs()) {
            System.out.println(xtp.getText());
          }
        } else if (sh instanceof XSLFTextShape) {
          XSLFTextShape shape = (XSLFTextShape) sh;
          // work with a shape that can hold text
        } else if (sh instanceof XSLFPictureShape) {
          XSLFPictureShape shape = (XSLFPictureShape) sh;
          // work with Picture
        }
      }
    }
  }

  //@Test
  public void inputPic() throws IOException, OpenXML4JException, XmlException {
    XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(tp1.toFile()));

    XSLFSlideMaster master = ppt.getSlideMasters().get(0);
    for (XSLFSlideLayout layout : master.getSlideLayouts()) {
      System.out.println("*****************" + layout.getType());
    }
    // title slide
    XSLFSlideLayout titleLayout = master.getLayout(SlideLayout.TITLE);
    XSLFSlideLayout layout1 = master.getLayout(SlideLayout.PIC_TX);

    XSLFSlide slide1 = ppt.createSlide(titleLayout);

    XSLFSlideLayout contentLayout = master.getLayout(SlideLayout.BLANK);

    XSLFSlide slide2 = ppt.createSlide(contentLayout);
    XSLFTextShape[] ph1 = slide2.getPlaceholders();

    XSLFTextShape titlePlaceholder1 = ph1[0];
    titlePlaceholder1.setText("left");
    XSLFTextShape right = ph1[1];
    right.setText("right");

    XSLFSlide slide3 = ppt.createSlide(layout1);
    /**
     XSLFTextShape[] ph1 = slide1.getPlaceholders();
     XSLFTextShape titlePlaceholder1 = ph1[0];
     //titlePlaceholder1.setText("This is a picture of an alarm clock");
     slide1.removeShape(titlePlaceholder1);
     XSLFTextShape subtitlePlaceholder1 = ph1[1];
     slide1.removeShape(subtitlePlaceholder1);
     XSLFTextShape thirdBlock = ph1[2];
     thirdBlock.setText("This may well be a caption");


     byte[] data;

     FileInputStream fis = new FileInputStream(path.toFile());
     data = IOUtils.toByteArray(fis);

     PictureData pictureIndex = ppt.addPicture(data, XSLFPictureData.PictureType.PNG);

     XSLFPictureShape shape = slide1.createPicture(pictureIndex);
     java.util.Date today = new java.util.Date();
     //subtitlePlaceholder1.setText(caption);
     thirdBlock.setText("hello simon");
     */

    FileOutputStream pptOutput = new FileOutputStream(outputPath.toFile());
    ppt.write(pptOutput);
    pptOutput.close();
    //fis.close();

  }

  public static byte[] toByteArray(int iSource, int iArrayLen) {
    byte[] bLocalArr = new byte[iArrayLen];
    for (int i = 0; (i < 4) && (i < iArrayLen); i++) {
      bLocalArr[i] = (byte) (iSource >> 8 * i & 0xFF);
    }
    return bLocalArr;
  }

  public XSLFSlide getCover() throws IOException {
    InputStream is = new FileInputStream(tp1.toFile());
    XMLSlideShow pptx = new XMLSlideShow(is);

    return pptx.getSlides().get(0);
  }

  //对pptx处理
  //@Test
  public void getPPT2007() throws IOException {

    InputStream is = new FileInputStream(tp1.toFile());
    XMLSlideShow pptx = new XMLSlideShow(is);

    String keywords = "";
    String summary = "";
    String title = "";

    List<XSLFSlide> slides = pptx.getSlides();
    System.out.println("ppt张数：" + slides.size());
    if (slides.size() > 0) {
      XSLFTextShape[] textshapes = slides.get(0).getPlaceholders();
      String title1 = slides.get(0).getTitle();

      for (int j = 0; j < textshapes.length; ++j) {
        Placeholder placeholder = textshapes[j].getTextType();
        System.out.println("页面1的占位类型 " + placeholder.name());
        if (placeholder == Placeholder.CENTERED_TITLE) {
          System.out.println("页面一的标题 " + textshapes[j].getText());
          title = textshapes[j].getText();
        }

      }
    }
    int i = 0;
    for (XSLFSlide slide : slides) {
      System.out.println("ppt" + (++i) + ":" + slide.getTitle());

      //          XSLFShape[] shapes = slide.getShapes();
      //          for(XSLFShape shape:shapes){
      //
      //              System.out.println("aizi "+shape.getShapeType());
      //
      //          }
      for (int j = 0; j < slide.getPlaceholders().length; ++j) {
        System.out.println("wenzi " + slide.getPlaceholder(j).getTextType().name());
      }
      System.out.println("****************************");
    }

    System.out.println("标题：" + title);
    System.out.println("关键词：" + keywords);
    System.out.println("摘要：" + summary);
    System.out.println(slides.get(0).getTitle());
    //      return list;
  }

  // 将byte数组bRefArr转为一个整数,字节数组的低位是整型的低字节位
  public static int toInt(byte[] bRefArr) {
    int iOutcome = 0;
    byte bLoop;

    for (int i = 0; i < bRefArr.length; i++) {
      bLoop = bRefArr[i];
      iOutcome += (bLoop & 0xFF) << (8 * i);
    }
    return iOutcome;
  }

  /**

   public  void outputPic() {

   InputStream is = new InputStream(new File(templatePath));
   // 加载PPT
   HSLFSlideShow _hslf = new HSLFSlideShow(templatePath);
   SlideShow _slideShow = new SlideShow(_hslf);

   // 获取PPT文件中的图片数据
   PictureData[] _pictures = _slideShow.getPictureData();

   // 循环读取图片数据
   for (int i = 0; i < _pictures.length; i++) {
   StringBuilder fileName = new StringBuilder(path);
   PictureData pic_data = _pictures[i];
   fileName.append(i);
   // 设置格式
   switch (pic_data.getType()) {
   case Picture.JPEG:
   fileName.append(".jpg");
   break;
   case Picture.PNG:
   fileName.append(".png");
   break;
   default:
   fileName.append(".data");
   }
   // 输出文件
   FileOutputStream fileOut = new FileOutputStream(new File(fileName.toString()));
   fileOut.write(pic_data.getData());
   fileOut.close();
   }

   }
   */

}
