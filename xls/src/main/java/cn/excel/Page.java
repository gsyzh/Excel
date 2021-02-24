package cn.excel;

import java.util.List;

public class Page {

    private String sheetName;

    /**
     * 页面遍历的数据 List 的泛型自行设置，如果所有数据都来着同一个类就写那个类，
     * 不是同一个类有继承就写继承类的泛型，没有就写问号。
     */
    private List<?> data;

    public Page(String sheetName,List<?> data) {
        super();
        this.sheetName = sheetName;
        this.data = data;
    }

    public Page() {
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<?> getData() {
        return data;
    }

    public void setData(List<?> data) {
        this.data = data;
    }
}
