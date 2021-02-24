package cn.excel.demo;

import cn.excel.Excel2Pdf;

public class Excel2PdfDemo {


    public static void main(String[] args) throws Exception {
        String path1 = "D:/excel及pdf导出文件/test.xlsx";
        //可以转换多个excel
        String[] sourcePath = {path1};
        String target = "D:/TEST/test1.pdf";
        //表头行数(由于是一个单元格一个单元格转换成pdfTableCell,表头不需要画上边框，需要判断表头行数判断)
        int headRows = 1;
        //每个pdfTbale下边距空间
        float spacingAfter = 30;
        Excel2Pdf.excuteExcel2Pdf(sourcePath,target,headRows,spacingAfter,"A3");

    }
}
