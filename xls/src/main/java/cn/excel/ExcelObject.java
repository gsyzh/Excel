package cn.excel;

import java.io.InputStream;

/**
 * @Description:ExcelObject对象类
 * @Auther: gsyzh
 * @Date: 2020-5-27 9:50
 */
public class ExcelObject {
    /**
     * 锚名称
     */
    private String anchorName;
    /**
     * ExcelSheet Stream
     */
    private InputStream inputStream;
    /**
     * POI ExcelSheet
     */
    private ExcelSheet excelSheet;

    public ExcelObject(InputStream inputStream){
        this.inputStream = inputStream;
        this.excelSheet = new ExcelSheet(this.inputStream);
    }

    public ExcelObject(String anchorName , InputStream inputStream){
        this.anchorName = anchorName;
        this.inputStream = inputStream;
        this.excelSheet = new ExcelSheet(this.inputStream);
    }

    public ExcelObject(InputStream inputStream,int index){
        this.anchorName = anchorName;
        this.inputStream = inputStream;
        this.excelSheet = new ExcelSheet(this.inputStream,index);
    }

    public ExcelObject(String anchorName , InputStream inputStream,int index){
        this.anchorName = anchorName;
        this.inputStream = inputStream;
        this.excelSheet = new ExcelSheet(this.inputStream);
    }
    public String getAnchorName() {
        return anchorName;
    }
    public void setAnchorName(String anchorName) {
        this.anchorName = anchorName;
    }
    public InputStream getInputStream() {
        return this.inputStream;
    }
    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }
    ExcelSheet getExcelSheet() {
        return excelSheet;
    }
}