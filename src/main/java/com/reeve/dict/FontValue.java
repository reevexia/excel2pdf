package com.reeve.dict;

/**
 * @Author Reeve
 * @Date 2021/8/13 10:39
 */
public enum  FontValue {

    SONGTI("宋体"),
    HEITI("黑体"),
    ADCVHC39B("AdvHC39b");

    private String FontName;

    FontValue(String fontName) {
        FontName = fontName;
    }

    public String getFontName() {
        return FontName;
    }
}
