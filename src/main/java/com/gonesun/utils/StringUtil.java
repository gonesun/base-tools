package com.gonesun.utils;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class StringUtil {


    public static boolean isNullOrEmpty(String str) {
        if (str == null) {
            return true;
        }
        return str.isEmpty();
    }

    /**
     * 判断字符串是否为空或者空串
     *
     * @param str
     * @return
     */
    public static boolean isEmpty(String str) {
        if (str == null || "null".equals(str) || "".equals(str.trim())) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * 日期信息转换 yyyy-MM-dd
     */
    public static String getFormatDateS(Date standardDate) {
        String tmp = "";
        try {
            if (standardDate != null) {
                DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                tmp = dateFormat.format(standardDate);
            }
        } catch (Exception ex) {
//            logger.error("getFormatDateS报错：", ex);
        }
        return tmp;
    }
}
