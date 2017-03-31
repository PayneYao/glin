package com.lozic.genpptx;

import com.lozic.genpptx.util.JProperties;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;

import static java.lang.System.exit;

/**
 * Created by wangqingwu on 16/12/13.
 * Since 16/12/13
 * Author Simon Gaius
 */
public class Main {
    private static Path confPath= Paths.get("d:\\gaius\\conf\\conf.properties");
    private static String projectName = "";
    private static String root="d:\\gaius\\projects\\";


    public static void main(String[] args) throws IOException, NoSuchFieldException {
        Properties p = new Properties();



        if(args.length==6){
            if("-conf".equals(args[0])){
               confPath = Paths.get(args[1]);
                p = JProperties.loadProperties(confPath.toString(), JProperties.BY_PROPERTIES);
                projectName = p.getProperty("excel.project");
                System.out.println("project name:"+projectName);
                root = p.getProperty("root");
                System.out.println("root:"+root);
                PptGen pptg = new PptGen(p,root,projectName);
                if("-cmd".equals(args[2])){
                    if("init".equals(args[3])){
                        //create dir
                        System.out.println("prepare for creating dir...");
                        pptg.prepareDirectory(root,projectName);
                    }else if("genmodel".equals(args[3])){

                        //create model

                        if("-template".equals(args[4])){
                            int model = Integer.parseInt(args[5]);
                            //create model according by model no.
                            switch(model){
                                case 1:
                                    //kfc-北京
                                    pptg.createConfig(p,root,projectName);
                                    break;
                                case 2:
                                    //广本北京
                                    pptg.createConfig2(p,root,projectName);
                                    break;
                                case 3:
                                    pptg.createConfig47(p,root,projectName,model);
                                    break;
                                case 4:
                                    //新励成
                                    pptg.createConfig47(p,root,projectName,model);
                                    break;
                                case 5:
                                    //搜狗模版
                                    pptg.createConfig5(p,root,projectName);
                                    break;
                                case 6:
                                    //吉利
                                    pptg.createConfig6(p,root,projectName);
                                    break;
                                case 7:
                                    //监测模版
                                    pptg.createConfig47(p,root,projectName,model);
                                    break;
                                case 8:
                                    //上版
                                    pptg.createConfig8(p,root,projectName);
                                    break;
                                default:
                                    System.out.println("template is not valid,please input 1-8");
                                    break;
                            }

                        }
                    }else if("genppt".equals(args[3])){
                        //create ppt
                        pptg.operImg2011();

                    }else{
                        System.out.println("missing command");
                    }
                }
            }


        }else{
            System.out.println("argument is not right");
            exit(0);
        }
        System.out.println(args[0]);
        System.out.println(args[1]);

    }
}
