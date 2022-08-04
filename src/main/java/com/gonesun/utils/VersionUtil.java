package com.gonesun.utils;

import java.util.ResourceBundle;

public class VersionUtil {

    private static String toolVersion = null;

    /**
     * 读取版本号
     */
    static {
        try {
            ResourceBundle resource = ResourceBundle.getBundle("tool");
            if (resource != null) {
                toolVersion = resource.getString("toolVersion");
            }
        } catch (Exception e) {
            toolVersion = "";
        }
    }

    public static String getToolVersion() {
        return toolVersion;
    }
}
