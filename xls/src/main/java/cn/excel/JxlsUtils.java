package cn.excel;

import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

/**
 * JxlsUtils工具类
 * @Description:通过定义excel模板，导出excel数据
 * @author gsyzh
 *
 */
public class JxlsUtils {

    /**
     * @Description: PdfTableExcel构造函数
     * @param templatePathIn（excel模板），templatePathOut（要导出的excel文件），model（导出的内容）
     */
	public static void exportExcel(String templatePathIn, String templatePathOut, Map<String, Object> model) throws IOException{
        FileOutputStream os = new FileOutputStream(templatePathOut);
        File template = getTemplateByPath(templatePathIn);
        if(template!=null){
            Context context = new Context();
            if (model != null) {
                for (String key : model.keySet()) {
                    context.putVar(key, model.get(key));
                }
            }
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();
            Transformer transformer  = jxlsHelper.createTransformer(new FileInputStream(template), os);
            jxlsHelper.processTemplate(context, transformer);
        }
        os.close();
	}

    /**
     * @Description:  获取jxls模版文件
     * @param path
     * @return File
     */
    public static File getTemplateByPath(String path){
        File template = new File(path);
        if(template.exists()){
            return template;
        }
        return null;
    }

    /**
     * @Description: 日期格式化
     * @param date,fmt
     * @return String
     */
    public static String dateFmt(Date date, String fmt) {
        if (date == null) {
            return "";
        }
        try {
            SimpleDateFormat dateFmt = new SimpleDateFormat(fmt);
            return dateFmt.format(date);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

}
