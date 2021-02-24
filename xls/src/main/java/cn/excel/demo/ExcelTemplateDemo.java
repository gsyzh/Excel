package cn.excel.demo;

import cn.excel.JxlsUtils;
import cn.excel.Page;
import cn.excel.demo.bean.*;

import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.*;

/**
 * Object collection output demo
 * @author gsyzh
 */
public class ExcelTemplateDemo {
    public static void main(String[] args) throws ParseException, IOException {
        List<Electric> electrics = generateSampleElectricData();
        Map<String, Object> model = new HashMap<String, Object>();
        model.put("electrics", electrics);
        //xls、xlsx都支持
        JxlsUtils.exportExcel("D:/TEST.xlsx", "D:/TESTRES.xlsx", model);
    }
    public static List<Electric> generateSampleElectricData() throws ParseException {
        List<Electric> electrics = new ArrayList<Electric>();
        electrics.add( new Electric("sheet1", new BigDecimal(10),  new BigDecimal(8),new BigDecimal(0.01),new BigDecimal(100), new BigDecimal(1),1));
        electrics.add( new Electric("sheet2", new BigDecimal(20),  new BigDecimal(28),new BigDecimal(0.02),new BigDecimal(200), new BigDecimal(2),2));
        electrics.add( new Electric("sheet3", new BigDecimal(30),  new BigDecimal(38),new BigDecimal(0.03),new BigDecimal(300), new BigDecimal(3),3));
        return electrics;
    }
}
