package com.reeve.pdf;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.pdf.BaseFont;
import com.reeve.dict.FontValue;
import org.apache.poi.hssf.usermodel.HSSFFont;

import java.io.IOException;

/**
 * @Author Reeve
 * @Date 2021/8/10 17:12
 */
public class PdfFont {

    public static Font getFont(HSSFFont hssfFont){
        String fontName = hssfFont.getFontName();
        boolean bold = hssfFont.getBold();
        short fontSize = (short) (hssfFont.getFontHeightInPoints() * 0.81);
        if(FontValue.SONGTI.getFontName().equals(fontName)){
            return SongFont(fontSize,bold);
        }
        if(FontValue.ADCVHC39B.getFontName().equals(fontName)){
            return ADCVHC39BFont(fontSize);
        }
        if(FontValue.HEITI.getFontName().equals(fontName)){
            return HeiFont(fontSize,bold);
        }
        return SongFont((short) 7,false);
    }

    /**
     * 默认字体
     * @param fontSzie
     * @param bold
     * @return
     */
    private static Font ChineseFont(short fontSzie,boolean bold) {
        BaseFont bf = null;
        Font font = null;
        try {
            bf = BaseFont.createFont("STSong-Light", "UniGB-UCS2-H",BaseFont.NOT_EMBEDDED);
            font = new Font(bf, fontSzie, Font.NORMAL);
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (bold) font.setStyle(Font.BOLD);
        return font;
    }

    /**
     * 条形码字体
     * @param fontSzie
     * @return
     */
    private static Font ADCVHC39BFont(short fontSzie) {
        Font font = null;
        String path = "src/main/resources/ADVHC39B.TTF";//自己的字体资源路径
        font = FontFactory.getFont(path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED,fontSzie, Font.NORMAL, BaseColor.BLACK);
        return font;
    }

    /**
     * 宋体
     * @param fontSzie
     * @param bold
     * @return
     */
    private static Font SongFont(short fontSzie, boolean bold) {
        Font font = null;
        String path = "src/main/resources/song.ttf";//自己的字体资源路径
        font = FontFactory.getFont(path, BaseFont.IDENTITY_H, BaseFont.EMBEDDED,fontSzie, Font.NORMAL, BaseColor.BLACK);
        if (bold) font.setStyle(Font.BOLD);
        return font;
    }

    /**
     * 黑体
     * @param fontSzie
     * @param bold
     * @return
     */
    private static Font HeiFont(short fontSzie, boolean bold) {
        Font font = null;
        String path = "src/main/resources/simhei.ttf";//自己的字体资源路径
        font = FontFactory.getFont(path, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED,fontSzie, Font.NORMAL, BaseColor.BLACK);
        if (bold) font.setStyle(Font.BOLD);
        return font;
    }
}
