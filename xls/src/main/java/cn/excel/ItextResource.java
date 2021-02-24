package cn.excel;

import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.pdf.BaseFont;
import org.apache.poi.hssf.usermodel.HSSFFont;

/**
 * @Description:iText资源类
 * @Auther: gsyzh
 * @Date: 2020-5-27 9:50
 */
public class ItextResource {
    /**
     * 中文字体支持
     */
    protected static BaseFont BASE_FONT_CHINESE;
    static {
        try {
            BASE_FONT_CHINESE = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
            // 搜尋系統,載入系統內的字型(慢)
            FontFactory.registerDirectories();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 將 POI Font 轉換到 iText Font
     * @param font
     * @return
     */
    public static Font getFont(HSSFFont font) {
        try {
            Font iTextFont = FontFactory.getFont(font.getFontName(),
                    BaseFont.IDENTITY_H, BaseFont.EMBEDDED,
                    font.getFontHeightInPoints());
            return iTextFont;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}