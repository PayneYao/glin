package com.lozic.genpptx.util;


import java.io.*;
import java.util.*;

/**
 * Created by wangqingwu on 16/12/11.
 * Since 16/12/11
 * Author Simon Gaius
 */

public class JProperties {


    public final static int BY_PROPERTIES = 1;

    public final static int BY_RESOURCEBUNDLE = 2;

    public final static int BY_PROPERTYRESOURCEBUNDLE = 3;

    public final static int BY_CLASS = 4;

    public final static int BY_CLASSLOADER = 5;

    public final static int BY_SYSTEM_CLASSLOADER = 6;


    public final static Properties loadProperties(final String name, final int type) throws IOException {

        Properties p = new Properties();

        InputStream in = null;
        InputStreamReader inr = null;

        if (type == BY_PROPERTIES) {

            inr = new InputStreamReader(new BufferedInputStream(new FileInputStream(name)),"UTF-8");

            assert (inr != null);

            p.load(inr);

        } else if (type == BY_RESOURCEBUNDLE) {

            ResourceBundle rb = ResourceBundle.getBundle(name, Locale.getDefault());

            assert (rb != null);

            p = new ResourceBundleAdapter(rb);

        } else if (type == BY_PROPERTYRESOURCEBUNDLE) {

            in = new BufferedInputStream(new FileInputStream(name));

            assert (in != null);

            ResourceBundle rb = new PropertyResourceBundle(in);

            p = new ResourceBundleAdapter(rb);

        } else if (type == BY_CLASS) {

            assert (JProperties.class.equals(new JProperties().getClass()));

            in = JProperties.class.getResourceAsStream(name);

            assert (in != null);

            p.load(in);

            //        return new JProperties().getClass().getResourceAsStream(name);

        } else if (type == BY_CLASSLOADER) {

            assert (JProperties.class.getClassLoader().equals(new JProperties().getClass().getClassLoader()));

            in = JProperties.class.getClassLoader().getResourceAsStream(name);

            assert (in != null);

            p.load(in);

            //       return new JProperties().getClass().getClassLoader().getResourceAsStream(name);

        } else if (type == BY_SYSTEM_CLASSLOADER) {

            in = ClassLoader.getSystemResourceAsStream(name);

            assert (in != null);

            p.load(in);

        }


        if (in != null) {

            in.close();

        }

        return p;


    }

    public static class ResourceBundleAdapter extends Properties {

        public ResourceBundleAdapter(ResourceBundle rb) {

            assert (rb instanceof java.util.PropertyResourceBundle);

            this.rb = rb;

            java.util.Enumeration e = rb.getKeys();

            while (e.hasMoreElements()) {

                Object o = e.nextElement();

                this.put(o, rb.getObject((String) o));

            }

        }


        private ResourceBundle rb = null;


        public ResourceBundle getBundle(String baseName) {

            return ResourceBundle.getBundle(baseName);

        }


        public ResourceBundle getBundle(String baseName, Locale locale) {

            return ResourceBundle.getBundle(baseName, locale);

        }


        public ResourceBundle getBundle(String baseName, Locale locale, ClassLoader loader) {

            return ResourceBundle.getBundle(baseName, locale, loader);

        }


        public Enumeration getKeys() {

            return rb.getKeys();

        }


        public Locale getLocale() {

            return rb.getLocale();

        }


        public Object getObject(String key) {

            return rb.getObject(key);

        }


        public String getString(String key) {

            return rb.getString(key);

        }


        public String[] getStringArray(String key) {

            return rb.getStringArray(key);

        }


        protected Object handleGetObject(String key) {

            return ((PropertyResourceBundle) rb).handleGetObject(key);

        }


    }
}
