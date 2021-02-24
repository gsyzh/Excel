package cn.excel;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.net.MalformedURLException;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * @Description: 将excel转成pdfTable
 * @Auther: gsyzh
 * @Date: 2020-5-27 9:50
 */
public class PdfTableExcel {
    protected ExcelObject excelObject;
    protected ExcelSheet excelSheet;
    protected boolean setting = false;

    /**
     * @Description: PdfTableExcel构造函数
     * @param excelObject
     */
    public PdfTableExcel(ExcelObject excelObject){
        this.excelObject = excelObject;
        this.excelSheet = excelObject.getExcelSheet();
    }

    /**
     * @Description: 获取Excel内容Table
     * @return PdfPTable
     * @throws BadElementException
     * @throws MalformedURLException
     * @throws IOException
     */
    public PdfPTable getTable(Document document,int headRows,float spacingAfter) throws BadElementException, MalformedURLException, IOException {
        Sheet sheet = this.excelSheet.getSheet();
        return toParseContent(sheet,document,headRows, spacingAfter);
    }

    /**
     * @Description: 解析excel内容
     * @throws BadElementException
     * @throws MalformedURLException
     * @throws IOException
     */
    protected PdfPTable toParseContent(Sheet sheet,Document document,int headRows, float spacingAfter) throws BadElementException, MalformedURLException, IOException{
        int rows = sheet.getPhysicalNumberOfRows();
        if (rows == 0){
            try {
                throw new Exception("表不可为空");
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        List<PdfPCell> cells = new ArrayList<PdfPCell>();
        float[] widths = null;
        float mw = 0;
        for (int i = 0; i < rows; i++) {
            Row row = sheet.getRow(i);
            int columns = row.getLastCellNum();

            float[] cws = new float[columns];
            for (int j = 0; j < columns; j++) {
                Cell cell = row.getCell(j);
                CellStyle cellStyle = cell.getCellStyle();
                short borderBottom = cellStyle.getBorderBottom();
                short borderLeft = cellStyle.getBorderLeft();

                if(cell.getCellType()==0){
//                    DecimalFormat df = new DecimalFormat("0.00");
//                    String getNumericCellValueStr = df.format(cell.getNumericCellValue());
                    NumberFormat nf = NumberFormat.getInstance();
                    cell.setCellValue(nf.format(cell.getNumericCellValue()));
                }

                if (cell == null) cell = row.createCell(j);

                float cw = getPOIColumnWidth(cell);
                cws[cell.getColumnIndex()] = cw;

                if(isUsed(cell.getColumnIndex(), row.getRowNum())){
                    continue;
                }

                cell.setCellType(Cell.CELL_TYPE_STRING);
                CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());

                int rowspan = 1;
                int colspan = 1;
                if (range != null) {
                    rowspan = range.getLastRow() - range.getFirstRow() + 1;
                    colspan = range.getLastColumn() - range.getFirstColumn() + 1;
                }

                PdfPCell pdfpCell = new PdfPCell();
                pdfpCell.setBackgroundColor(new BaseColor(POIUtil.getRGB(cell.getCellStyle().getFillForegroundColorColor())));
                pdfpCell.setColspan(colspan);
                pdfpCell.setRowspan(rowspan);
                pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment()));
                pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment()));
                pdfpCell.setPhrase(getPhrase(cell));
                //pdfpCell.setBorder(14、15);全边
                //pdfpCell.setBorder(2);有下边
//                if (i < headRows-1){
//                    pdfpCell.setBorder(0);
//                }else if (i == headRows-1){
//                    pdfpCell.setBorder(2);
//                }else{
//                    pdfpCell.setBorder(15);
//                }

                if (borderBottom==0){
                    if (borderLeft == 0){
                        pdfpCell.setBorder(0);
                    }else{
                        pdfpCell.setBorder(15);
                    }
                }

                if (sheet.getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
                    pdfpCell.setFixedHeight(this.getPixelHeight(row.getHeightInPoints()));
                }

                addBorderByExcel(pdfpCell, cell.getCellStyle());
                addImageByPOICell(pdfpCell , cell , cw);

                cells.add(pdfpCell);
                j += colspan - 1;
            }

            float rw = 0;
            for (int j = 0; j < cws.length; j++) {
                rw += cws[j];
            }
            if (rw > mw ||  mw == 0) {
                widths = cws;
                mw = rw;
            }
        }

        PdfPTable table = new PdfPTable(widths);

        if (document.getPageSize().getWidth() < mw){
            table.setWidthPercentage(100);
        }else{
            table.setTotalWidth(mw);
            table.setLockedWidth(true);
        }

        for (PdfPCell pdfpCell : cells) {
            table.addCell(pdfpCell);
        }
        table.setSpacingAfter(spacingAfter);
        return table;
    }

    /**
     * @Description: 获取锚
     * @param cell
     * @return Phrase
     */
    protected Phrase getPhrase(Cell cell) {
        if(this.setting || this.excelObject.getAnchorName() == null){
            return new Phrase(cell.getStringCellValue(), getFontByExcel(cell.getCellStyle()));
        }
        Anchor anchor = new Anchor(cell.getStringCellValue() , getFontByExcel(cell.getCellStyle()));
        anchor.setName(this.excelObject.getAnchorName());
        this.setting = true;
        return anchor;
    }

    /**
     * @Description: 加excel图片
     * @param cell
     */
    protected void addImageByPOICell(PdfPCell pdfpCell , Cell cell , float cellWidth) throws BadElementException, MalformedURLException, IOException{
        POIImage poiImage = new POIImage().getCellImage(cell);
        byte[] bytes = poiImage.getBytes();
        if(bytes != null){
//           double cw = cellWidth;
//           double ch = pdfpCell.getFixedHeight();
//
//           double iw = poiImage.getDimension().getWidth();
//           double ih = poiImage.getDimension().getHeight();
//
//           double scale = cw / ch;
//
//           double nw = iw * scale;
//           double nh = ih - (iw - nw);
//
//           POIUtil.scale(bytes , nw  , nh);
            pdfpCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            Image image = Image.getInstance(bytes);
            pdfpCell.setImage(image);
        }
    }

    /**
     * @Description: 处理像素高度
     * @param poiHeight
     * @return 像素
     */
    protected float getPixelHeight(float poiHeight){
        float pixel = poiHeight / 28.6f * 26f;
        return pixel;
    }

    /**
     * @Description: 此处获取Excel的列宽像素(无法精确实现,期待有能力的朋友进行改善此处)
     * @param cell
     * @return 像素宽
     */
    protected int getPOIColumnWidth(Cell cell) {
        int poiCWidth = excelSheet.getSheet().getColumnWidth(cell.getColumnIndex());
        int colWidthpoi = poiCWidth;
        int widthPixel = 0;
        if (colWidthpoi >= 416) {
            widthPixel = (int) (((colWidthpoi - 416.0) / 256.0) * 8.0 + 13.0 + 0.5);
        } else {
            widthPixel = (int) (colWidthpoi / 416.0 * 13.0 + 0.5);
        }
        return widthPixel;
    }

    /**
     * @Description: Excel行
     * @param rowIndex,colIndex
     * @return CellRangeAddress
     */
    protected CellRangeAddress getColspanRowspanByExcel(int rowIndex, int colIndex) {
        CellRangeAddress result = null;
        Sheet sheet = excelSheet.getSheet();
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstColumn() == colIndex && range.getFirstRow() == rowIndex) {
                result = range;
            }
        }
        return result;
    }

    /**
     * @Description: 单元格是否被使用过
     * @param colIndex,rowIndex
     * @return Font
     */
    protected boolean isUsed(int colIndex , int rowIndex){
        boolean result = false;
        Sheet sheet = excelSheet.getSheet();
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            if (firstRow < rowIndex && lastRow >= rowIndex) {
                if(firstColumn <= colIndex && lastColumn >= colIndex){
                    result = true;
                }
            }
        }
        return result;
    }

    /**
     * @Description: excel字体
     * @param style
     * @return Font
     */
    protected Font getFontByExcel(CellStyle style) {
        //pdf font
        Font result = new Font(ItextResource.BASE_FONT_CHINESE , 8 , Font.NORMAL);
        Workbook wb = excelSheet.getWorkbook();

        short index = style.getFontIndex();
        //excel font
        org.apache.poi.ss.usermodel.Font font = wb.getFontAt(index);
        short fontSize = font.getFontHeightInPoints();
        if (fontSize > 13){//excel字体大小
            //13是pdf最大字体了
            result.setSize(13);
        }else{
            //10、11、12不存在
            result.setSize(8);
        }
        if(font.getBoldweight() == org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD){
            result.setStyle(Font.BOLD);
        }

        HSSFColor color = HSSFColor.getIndexHash().get(font.getColor());

        if(color != null){
            int rbg = POIUtil.getRGB(color);
            result.setColor(new BaseColor(rbg));
        }

        FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
        if(underline == FontUnderline.SINGLE){
            String ulString = Font.FontStyle.UNDERLINE.getValue();
            result.setStyle(ulString);
        }
        return result;
    }

    /**
     * @Description: 边框样式
     * @param cell，style
     */
    protected void addBorderByExcel(PdfPCell cell , CellStyle style) {
        Workbook wb = excelSheet.getWorkbook();
        cell.setBorderColorLeft(new BaseColor(POIUtil.getBorderRBG(wb,style.getLeftBorderColor())));
        cell.setBorderColorRight(new BaseColor(POIUtil.getBorderRBG(wb,style.getRightBorderColor())));
        cell.setBorderColorTop(new BaseColor(POIUtil.getBorderRBG(wb,style.getTopBorderColor())));
        cell.setBorderColorBottom(new BaseColor(POIUtil.getBorderRBG(wb,style.getBottomBorderColor())));
    }

    /**
     * @Description: 垂直对齐方式
     * @param align
     * @return 像素宽
     */
    protected int getVAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.VERTICAL_BOTTOM) {
            result = Element.ALIGN_BOTTOM;
        }
        if (align == CellStyle.VERTICAL_CENTER) {
            result = Element.ALIGN_MIDDLE;
        }
        if (align == CellStyle.VERTICAL_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.VERTICAL_TOP) {
            result = Element.ALIGN_TOP;
        }
        return result;
    }

    /**
     * @Description: 水平对齐方式
     * @param align
     * @return 像素宽
     */
    protected int getHAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.ALIGN_LEFT) {
            result = Element.ALIGN_LEFT;
        }
        if (align == CellStyle.ALIGN_RIGHT) {
            result = Element.ALIGN_RIGHT;
        }
        if (align == CellStyle.ALIGN_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.ALIGN_CENTER) {
            result = Element.ALIGN_CENTER;
        }
        return result;
    }
}