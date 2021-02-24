package cn.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * @Description: Excel sheet页
 * @Auther: gsyzh
 * @Date: 2020-5-27 9:50
 */
public class ExcelSheet {

    protected Workbook wb;
    protected Sheet sheet;

    public ExcelSheet(InputStream is, int index) {
        try {
            this.wb = WorkbookFactory.create(is);
            this.sheet = wb.getSheetAt(index);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public ExcelSheet(InputStream is) {
        try {
            this.wb = WorkbookFactory.create(is);
            this.sheet = wb.getSheetAt(wb.getActiveSheetIndex());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Workbook getWorkbook(){
        return wb;
    }
}
