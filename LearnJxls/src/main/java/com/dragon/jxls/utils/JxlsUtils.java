package com.dragon.jxls.utils;

import com.dragon.common.Constants;
import org.apache.log4j.Logger;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.transform.poi.PoiTransformer;
import org.jxls.util.JxlsHelper;

import java.io.*;
import java.util.*;

/**
 * 批注说明：
 * 1.XLS Area 是JxlsPlus中的一个重要概念，它代表excel模板中需要被解析的矩形区域，由A1到最后一个单元格表示，有利于加快解析速度。
 * 2.Jxls默认支持Apache JEXL表达式语言，用于在模板中操作Java对象的属性及方法，类似于EL表达式。
 * 版本号须对应上，否则会抛出Transformer异常
 *
 * 例子: https://www.cnblogs.com/dw3306/p/11098841.html
 * https://blog.csdn.net/dai200702008/article/details/84847315?ops_request_misc=%257B%2522request%255Fid%2522%253A%2522160545757219724842922823%2522%252C%2522scm%2522%253A%252220140713.130102334..%2522%257D&request_id=160545757219724842922823&biz_id=0&utm_medium=distribute.pc_search_result.none-task-blog-2~all~top_click~default-1-84847315.first_rank_ecpm_v3_pc_rank_v2&utm_term=jxls&spm=1018.2118.3001.4449
 *
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 14:08
 * @description：
 * @modified By：
 * @version: $
 */
public class JxlsUtils {
    private static transient Logger logger = Logger.getLogger(JxlsUtils.class);

    // 初始加载
    static {
        XlsCommentAreaBuilder.addCommandMapping("merge", MergeCommand.class);// 拓展合并单元格方法
    }

    /**
     *
     * @param is
     * @param os
     * @param model
     * @throws IOException
     */
    public static void exportExcel(InputStream is, OutputStream os, Map<String, Object> model) throws IOException {
        Context context = PoiTransformer.createInitialContext();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }
        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
        Transformer transformer = jxlsHelper.createTransformer(is, os);
        // 获得配置
        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator) transformer.getTransformationConfig().getExpressionEvaluator();
        // 设置静默模式，不报警告
        // evaluator.getJexlEngine().setSilent(true);
        // 函数强制，自定义功能
        Map<String, Object> funcs = new HashMap<String, Object>();
        funcs.put("utils", new JxlsTools()); // 添加自定义功能
        evaluator.getJexlEngine().setFunctions(funcs);// .setFunctions(funcs);
        // 必须要这个，否者表格函数统计会错乱
        jxlsHelper.setUseFastFormulaProcessor(false).processTemplate(context, transformer);
    }

    /**
     * 生成文件通用替换方法1
     * @param templatePath:模板路径
     * @param createFilePath：生成文件路径
     * @param model:参数
     * @throws Exception
     */
    public static Map<String,Object> exportExcel(String templatePath, String createFilePath, Map<String, Object> model) {
        Map<String,Object> returnMap = new HashMap<>();
        returnMap.put(Constants.RETURN_CODE, Constants.RETURN_ERR_CODE);
        OutputStream os = null;
        InputStream is = null;
        File template = getTemplate(templatePath);
        if (template == null || !template.exists()) {
            logger.info("导出模板文件获取为空");
            throw new RuntimeException("模板文件未找到");
        }
        try {
            os = new FileOutputStream(createFilePath);
            is =  new FileInputStream(template);
            exportExcel(is, os, model);
            returnMap.put(Constants.RETURN_CODE, Constants.RETURN_SUC_CODE);
        } catch (FileNotFoundException e) {
            logger.error(e.getMessage(),e);
            returnMap.put(Constants.RETURN_MSG, e.toString());
        } catch (IOException e) {
            logger.error(e.getMessage(),e);
            returnMap.put(Constants.RETURN_MSG, e.toString());
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            returnMap.put(Constants.RETURN_MSG, e.toString());
        }finally {
            if(os != null){
                try {
                    os.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
            }
            if(is != null){
                try {
                    is.close();
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
            }
        }
        return returnMap;
    }

    // 获取jxls模版文件
    public static File getTemplate(String path) {
        File template = new File(path);
        if (template.exists()) {
            return template;
        }
        return null;
    }
}
