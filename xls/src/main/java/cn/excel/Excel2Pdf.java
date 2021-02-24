package cn.excel;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.dom4j.DocumentException;

import java.io.*;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

/**
 * @Description: Excel转Pdf工具类
 * @Auther: gsyzh
 * @Date: 2020-5-27 9:50
 */
public class Excel2Pdf extends ItextPdf {
    protected List<ExcelObject> excelObjects = new ArrayList<ExcelObject>();

    /**
     * @Description: Excel转Pdf
     * @Auther: gsyzh
     * @Date: 2020-5-27 9:50
     */
    public static void excuteExcel2Pdf(String[] sourcePath, String targetPath, int headRows, float spacingAfter,String pageSize) {
        List<ExcelObject> excelObjects = new ArrayList<ExcelObject>();
        try {
            for (int i = 0 ; i < sourcePath.length; i++) {
                FileInputStream in = new FileInputStream(new File(sourcePath[i]));
                Workbook wb = WorkbookFactory.create(in);
                int counts = wb.getNumberOfSheets();
                for (int j = 0 ; j < counts ; j++){
                    //上次流被用过关闭了，需要再创建一次。
                    FileInputStream fis = new FileInputStream(new File(sourcePath[i]));
                    excelObjects.add(new ExcelObject(fis,j));
                    fis.close();
                }
                in.close();
            }
            FileOutputStream fos = new FileOutputStream(new File(targetPath));
            Excel2Pdf pdf = new Excel2Pdf(excelObjects, fos);
            pdf.convert(headRows,spacingAfter,pageSize);
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (com.itextpdf.text.DocumentException e) {
            e.printStackTrace();
        }

    }

    /**
     * @Description: 导出单项PDF，不包含目录
     * @param excelObject
     */
    public Excel2Pdf(ExcelObject excelObject , OutputStream os) {
        this.excelObjects.add(excelObject);
        this.os = os;
    }

    /**
     * @Description: 导出多项PDF，包含目录
     * @param excelObjects
     */
    public Excel2Pdf(List<ExcelObject> excelObjects , OutputStream os) {
        this.excelObjects = excelObjects;
        this.os = os;
    }

    /**
     * @Description: 转换调用
     * @param headRows（表头行数），spacingAfter（表格后预留空间）
     * @throws DocumentException
     * @throws MalformedURLException
     * @throws IOException
     */
    public void convert(int headRows,float spacingAfter,String pageSize) throws DocumentException, MalformedURLException, IOException, com.itextpdf.text.DocumentException {
        switch(pageSize){
            case "A0" :
                getDocument().setPageSize(PageSize.A0.rotate());
                break;
            case "A1" :
                getDocument().setPageSize(PageSize.A1.rotate());
                break;
            case "A2" :
                getDocument().setPageSize(PageSize.A2.rotate());
                break;
            case "A3" :
                getDocument().setPageSize(PageSize.A3.rotate());
                break;
            case "A4" :
                getDocument().setPageSize(PageSize.A4.rotate());
                break;
            case "A5" :
                getDocument().setPageSize(PageSize.A5.rotate());
                break;
            default :
                //语句
                getDocument().setPageSize(PageSize.A4.rotate());
        }
        PdfWriter writer = PdfWriter.getInstance(getDocument(), os);
        writer.setPageEvent(new PDFPageEvent());
        //Open document
        getDocument().open();
        //Single one
        if(this.excelObjects.size() <= 1){
            PdfPTable table = this.toCreatePdfTable(this.excelObjects.get(0) ,  getDocument() , writer, headRows, spacingAfter);
            getDocument().add(table);
        }
        //Multiple ones
        if(this.excelObjects.size() > 1){
            toCreateContentIndexes(writer , this.getDocument() , this.excelObjects);

            for (int i = 0; i < this.excelObjects.size(); i++) {
                PdfPTable table = this.toCreatePdfTable(this.excelObjects.get(i) , getDocument() , writer, headRows, spacingAfter);
                getDocument().add(table);
            }
        }
        getDocument().close();
    }

    /**
     * @Description: 创建PdfTable
     * @throws DocumentException
     * @throws MalformedURLException
     * @throws IOException
     */
    protected PdfPTable toCreatePdfTable(ExcelObject object , Document document , PdfWriter writer,int headRows,float spacingAfter) throws MalformedURLException, IOException, DocumentException, BadElementException {
        PdfPTable table = new PdfTableExcel(object).getTable(getDocument(),headRows, spacingAfter);
        table.setKeepTogether(true);
//      table.setWidthPercentage(new float[]{100} , writer.getPageSize());
        table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
        return table;
    }

    /**
     * @Description: 内容索引创建
     * @throws DocumentException
     */
    protected void toCreateContentIndexes(PdfWriter writer , Document document , List<ExcelObject> objects) throws DocumentException, com.itextpdf.text.DocumentException {
        PdfPTable table = new PdfPTable(1);
        table.setKeepTogether(true);
        table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
        //从资源中获取字体
        Font font = new Font(ItextResource.BASE_FONT_CHINESE , 12 , Font.NORMAL);
        font.setColor(new BaseColor(0,0,255));

        for (int i = 0; i < objects.size(); i++) {
            ExcelObject o = objects.get(i);
            String text = o.getAnchorName();
            Anchor anchor = new Anchor(text , font);
            anchor.setReference("#" + o.getAnchorName());
            PdfPCell cell = new PdfPCell(anchor);
            cell.setBorder(0);
            //
            table.addCell(cell);
        }
        document.add(table);
    }

    /**
     * @ClassName: PDFPageEvent
     * @Description: 事件 -> 页码控制
     * @Author: gsyzh
     */
    private static class PDFPageEvent extends PdfPageEventHelper {
        protected PdfTemplate template;
        public BaseFont baseFont;

        @Override
        public void onStartPage(PdfWriter writer, Document document) {
            try{
                this.template = writer.getDirectContent().createTemplate(100, 100);
                this.baseFont = new Font(ItextResource.BASE_FONT_CHINESE , 8, Font.NORMAL).getBaseFont();
            } catch(Exception e) {
                throw new ExceptionConverter(e);
            }
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            //在每页结束的时候把“第x页”信息写道模版指定位置
            PdfContentByte byteContent = writer.getDirectContent();
            String text = "第" + writer.getPageNumber() + "页";
            float textWidth = this.baseFont.getWidthPoint(text, 8);
            float realWidth = document.right() - textWidth;

            byteContent.beginText();
            byteContent.setFontAndSize(this.baseFont , 10);
            byteContent.setTextMatrix(realWidth , document.bottom());
            byteContent.showText(text);
            byteContent.endText();
            byteContent.addTemplate(this.template , realWidth , document.bottom());
        }
    }
}