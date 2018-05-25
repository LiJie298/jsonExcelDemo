package com.lijie.util;

import com.google.common.io.Resources;

import java.io.*;
import java.net.InetAddress;
import java.util.Properties;
import java.util.concurrent.locks.ReentrantLock;

/**
 * @author lijie7
 * @date 2018/1/4
 * @Description
 * @modified By
 */
public class ExcelConfig {
    private static final String fileName = "excelFile.properties";
    private static final ReentrantLock mainLock = new ReentrantLock();
    private static ExcelConfig me;
    private static Properties properties;
    public static String localIp;

    static {
        try {
            localIp = InetAddress.getLocalHost().getHostAddress();
        } catch (Exception e) {
            Log.error("初始化配置文件失败", e.getMessage());
        }
    }

    private ExcelConfig() {
        init();
    }

    public static ExcelConfig getMe() {
        mainLock.lock();
        try {
            return me == null ? me = new ExcelConfig() : me;
        } finally {
            mainLock.unlock();
        }
    }



    public static void init(String filePath) {
        mainLock.lock();
        InputStream stream = null;
        try {
            if (properties != null) {
                properties.clear();
            } else {
                properties = new Properties();
            }
            stream = new FileInputStream(filePath);
            BufferedReader bf = new BufferedReader(new InputStreamReader(stream));
            properties.load(bf);
            Log.info("配置文件加载成功...");
        } catch (FileNotFoundException e) {
            Log.error("properties file not found!!!　", e.getMessage());
            System.exit(-1);
        } catch (IOException e) {
            Log.error("读取配置文件出错", e.getMessage());
        } finally {
            mainLock.unlock();
            if (stream != null) {
                try {
                    stream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void init() {
        init(Resources.getResource(fileName).getPath());
    }

    public void put(String key, String val) {
        properties.put(key, val);
    }

    public Properties getAll() {
        return properties;
    }

    public String get(String key) {
        return properties.getProperty(key);
    }
}
