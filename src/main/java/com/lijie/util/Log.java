package com.lijie.util;


import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

/**
 * Created by lijie on 2017/11/10.
 */
public class Log {

    private static Logger logger = LogManager.getLogger("infoLog");
    private static Logger errorLogger = LogManager.getLogger("errorLog");

    /**
     * 记录info信息
     *
     * @param msg 需要记录的字符串
     */
    public static void info(String msg) {
        logger.info(Thread.currentThread().getName() + " | " + msg);
    }

    /**
     * 记录错误
     */
    public static void error(String title, String msg) {
        errorLogger.error("title:" + title + " | msg:" + msg);
    }

}
