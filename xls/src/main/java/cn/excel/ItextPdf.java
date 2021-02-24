package cn.excel;

import com.itextpdf.text.Document;

import java.io.OutputStream;

/**
 * @Description: ItextPdf工具类
 * @Auther: gsyzh
 * @Date: 2020-5-27 9:50
 */
public class ItextPdf {

    protected Document document;

    protected OutputStream os;

    public Document getDocument() {
        if (document == null) {
            document = new Document();
        }
        return document;
    }
}