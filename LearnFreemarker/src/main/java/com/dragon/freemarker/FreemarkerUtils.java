package com.dragon.freemarker;

import com.dragon.common.entity.UserEntity;
import freemarker.cache.StringTemplateLoader;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.StringWriter;
import java.util.HashMap;
import java.util.Map;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 11:02
 * @description：
 * @modified By：
 * @version: $
 */
public class FreemarkerUtils {
    private static final Logger logger = LoggerFactory.getLogger(FreemarkerUtils.class);

    private static Configuration configuration = null;
    /**
     * 模板名称
     */
    private static String templateName = "template";

    public static void main(String[] args) {
        // 默认模板1
        String templateValueOne = "尊敬${name}客户，您的（账号：${account}）于今日${thisTime}分发起一笔数额为${amount}元的${type}申请；";
        // 默认模板2
        String templateValueTwo = "Hello！本人姓名：${name}，性别：${sex}，年龄：${age}岁";
        try {
            // 方法一
            Map<String, Object> userMap = new HashMap<String,Object>();
            userMap.put("name","DragonWen");
            userMap.put("account","920310436");
            userMap.put("thisTime","13:15");
            userMap.put("amount","50000");
            userMap.put("type","转账");
            String valueStrOne = initFreemarkerByModel(templateValueOne, userMap);
            logger.info("生成模板内容一 = " + valueStrOne);

            // 方法二
            UserEntity userEntity = new UserEntity();
            userEntity.setName("DragonWen");
            userEntity.setAge("26");
            userEntity.setSex("男");
            String valueStrTwo = initFreemarkerByEntity(templateValueTwo, userEntity);
            logger.info("生成模板内容二 = " + valueStrTwo);
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
        }
    }

    /**
     * 使用模型
     * @param templateValue
     * @param pushModelMap
     * @return
     */
    public static String initFreemarkerByModel(String templateValue,Map<String, Object> pushModelMap) {
        if(configuration == null) {
            configuration = configuration();
        }
        String strValue ="";
        try {
            Configuration configuration = configuration();
            strValue = processTemplate(pushModelMap, configuration, templateName, templateValue);
            return strValue;
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            logger.error("模板异常="+e.getMessage());
        }
        return strValue;
    }

    /**
     *
     * @param templateValue
     * @param pushModel
     * @return
     */
    public static String initFreemarkerByEntity(String templateValue, Object pushModel) {
        if(configuration == null) {
            configuration = configuration();
        }
        String strValue ="";
        try {
            Configuration configuration = configuration();
            strValue = processModalTemplate(pushModel, configuration, templateName, templateValue);
            return strValue;
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
            logger.error("模板异常="+e.getMessage());
        }
        return strValue;
    }

    /**
     ** 解析模板
     * @param configuration
     * @param templateName
     * @throws IOException
     * @throws TemplateException
     */
    private static String processTemplate(Map<String, Object> pushModelMap, Configuration configuration, String templateName, String templateValue)
            throws IOException, TemplateException {
        StringWriter stringWriter = new StringWriter();
        Template template = new Template(templateName, templateValue, configuration);
        template.process(pushModelMap, stringWriter);
        return stringWriter.toString();
    }

    /**
     *
     * @param dataModel
     * @param configuration
     * @param templateName
     * @param templateValue
     * @return
     * @throws IOException
     * @throws TemplateException
     */
    private static String processModalTemplate(Object dataModel,Configuration configuration, String templateName, String templateValue)
            throws IOException, TemplateException {
        StringWriter stringWriter = new StringWriter();
        Template template = new Template(templateName, templateValue, configuration);
        template.process(dataModel, stringWriter);
        return stringWriter.toString();
    }

    /**
     *
     * 配置 freemarker configuration
     * @return
     */
    private static Configuration configuration() {
        Configuration configuration = new Configuration(Configuration.VERSION_2_3_23);
        StringTemplateLoader templateLoader = new StringTemplateLoader();
        configuration.setTemplateLoader(templateLoader);
        configuration.setDefaultEncoding("UTF-8");
        return configuration;
    }
}
