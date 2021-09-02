package com.reeve.util;

/**
 * @Author Reeve
 * @Date 2021/8/13 11:04
 */
public class StringUtil {

    public static boolean compare(String v1, String v2) {
        if (v1 == "" && v2 == "") return true;
        if (v1 == null && v2 == null) return true;
        if (v1=="" && v2 == null || v2=="" && v1 == null) return false;
        if (v1.length()==v2.length()){
            for (int i = 0; i < v1.length(); i++) {

            }
        }
        return false;
    }
}
