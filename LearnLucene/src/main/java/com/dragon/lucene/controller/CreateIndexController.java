package com.dragon.lucene.controller;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.dragon.common.Constants;
import com.dragon.lucene.utils.LuceneUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.lucene.analysis.Analyzer;
import org.apache.lucene.analysis.standard.StandardAnalyzer;
import org.apache.lucene.document.Document;
import org.apache.lucene.document.Field;
import org.apache.lucene.document.FieldType;
import org.apache.lucene.index.IndexOptions;
import org.apache.lucene.index.IndexWriter;
import org.apache.lucene.index.IndexWriterConfig;
import org.apache.lucene.store.Directory;
import org.apache.lucene.store.FSDirectory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 11:19
 * @description：创建索引Controller
 * @modified By：
 * @version: $
 */
public class CreateIndexController {
    private static final Logger logger = LoggerFactory.getLogger(CreateIndexController.class);

    public static void main(String[] args) {
        Map<String, Object> info = createIndexInfo();
        String code = LuceneUtils.getStringToMap(info, Constants.RETURN_CODE);
        String msg = LuceneUtils.getStringToMap(info, Constants.RETURN_MSG);
        logger.info("code:"+code+"=====msg:"+msg);
    }

    /**
     * 创建查询索引文件（初始化Lucene）
     * 在实际项目开发中Lucene初始化可以在Service使用@Async("asyncServiceExecutor")异步初始化
     * @return
     */
    private static Map<String, Object> createIndexInfo(){
        String messagePre = "索引文件生成createIndexInfo方法:";

        Map<String,Object> returnMap = new HashMap<>();
        returnMap.put(Constants.RETURN_CODE, Constants.RETURN_SUC_CODE);
        returnMap.put(Constants.RETURN_MSG, "执行成功");

        // 这里为方便演示自己造了点数据，实际项目中可从数据库查询数据
        List<Map<String,Object>> dataList = new ArrayList<>();
        String jsonData = readLocalFile();
        JSONObject obj = JSONObject.parseObject(jsonData);
        JSONArray dataArr = obj.getJSONArray("data");
        if(obj != null){
            for (Object data : dataArr) {
                JSONObject dataObj =JSONObject.parseObject(data.toString());
                String id = dataObj.get("id") == null ? "" : dataObj.get("id").toString();
                String companyName = dataObj.get("companyName") == null ? "" : dataObj.get("companyName").toString();
                String companyAddress = dataObj.get("companyAddress") == null ? "" : dataObj.get("companyAddress").toString();
                Map<String,Object> dataMap = new HashMap<>();
                dataMap.put("id",id);
                dataMap.put("companyName",companyName);
                dataMap.put("companyAddress",companyAddress);
                dataList.add(dataMap);
            }
        }

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        logger.info("{}-查询start:{}",messagePre,sdf.format(new Date()));

        if(CollectionUtils.isNotEmpty(dataList)){
            logger.info("{}索引创建开始start:{}",messagePre,sdf.format(new Date()));
            // 索引生成路径
            String filePath = LuceneUtils.getFilePath(Constants.TEST_CACHE_FILENAME);

            // 创建分词器对象
            Analyzer analyzer = new StandardAnalyzer();
            IndexWriterConfig icw = new IndexWriterConfig(analyzer);
            // CREATE 表示先清空索引再重新创建
            icw.setOpenMode(IndexWriterConfig.OpenMode.CREATE);
            Directory dir = null;
            IndexWriter inWriter = null;
            // 存储索引的目录
            Path indexPath = Paths.get(filePath);
            // 批量写入多个文档
            ArrayList<Document> documentList = new ArrayList<Document>();
            try {
                if (!Files.isReadable(indexPath)) {
                    logger.info(messagePre+"索引目录 '" + indexPath.toAbsolutePath() + "' 不存在或者不可读,请检查");
                    throw new RuntimeException(messagePre+"索引目录不存在或者不可读,请检查");
                }
                dir = FSDirectory.open(indexPath);
                inWriter = new IndexWriter(dir, icw);

                // 设置 字段 索引并存储
                FieldType idType = new FieldType();
                idType.setIndexOptions(IndexOptions.DOCS);
                idType.setStored(true);

                // 设置 字段（标题）索引文档、词项频率、位移信息和偏移量，存储并词条化
                FieldType titleType = new FieldType();
                titleType.setIndexOptions(IndexOptions.DOCS_AND_FREQS_AND_POSITIONS_AND_OFFSETS);
                titleType.setStored(true);
                titleType.setTokenized(true);

                // 设置内容 字段属性
                FieldType contentType = new FieldType();
                contentType.setIndexOptions(IndexOptions.DOCS_AND_FREQS_AND_POSITIONS_AND_OFFSETS);
                contentType.setStored(true);
                contentType.setTokenized(true);
                contentType.setStoreTermVectors(true);
                contentType.setStoreTermVectorPositions(true);
                contentType.setStoreTermVectorOffsets(true);

                // 如果要根据这个字段进行搜索，那么这个字段就必须创建索引
                for (Map<String, Object> dataMap : dataList) {
                    String id = LuceneUtils.getStringToMap(dataMap,"id");
                    String companyName = LuceneUtils.getStringToMap(dataMap,"companyName");
                    String companyAddress = LuceneUtils.getStringToMap(dataMap,"companyAddress");

                    Document doc = new Document();
                    // ID
                    doc.add(new Field("id", id, idType));
                    // 名称
                    doc.add(new Field("companyName", companyName, titleType));
                    // 地址
                    doc.add(new Field("companyAddress", companyAddress, contentType));
                    documentList.add(doc);
                    doc = null;
                }
                inWriter.addDocuments(documentList);
                inWriter.commit();
            } catch (IOException e) {
                logger.error("创建索引异常：",e);
                logger.error("创建索引异常："+e.getMessage());
                returnMap.put(Constants.RETURN_CODE, Constants.RETURN_ERR_CODE);
                returnMap.put(Constants.RETURN_MSG, "创建索引异常");
            }finally{
                try {
                    if(inWriter != null){
                        inWriter.close();
                    }
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
                try {
                    if(dir != null){
                        dir.close();
                    }
                } catch (IOException e) {
                    logger.error(e.getMessage(),e);
                }
                if(documentList != null){
                    documentList = null;
                }
            }
            logger.info("{}索引创建结束end:{}",messagePre,sdf.format(new Date()));
        }else {
            throw new RuntimeException();
        }
        logger.info("{}-查询end:{}",messagePre,sdf.format(new Date()));
        return returnMap;
    }

    /**
     * 读取本地文件内容
     * @return
     */
    private static String readLocalFile() {
        BufferedReader reader = null;
        StringBuffer sbf = new StringBuffer();
        try {
            File file = new File("D:\\Project\\idea-workspace\\LearnJava\\LearnLucene\\src\\main\\resources\\data.txt");
            reader = new BufferedReader(new FileReader(file));
            String tempStr;
            while ((tempStr = reader.readLine()) != null) {
                sbf.append(tempStr);
            }
            reader.close();
            return sbf.toString();
        } catch (IOException e) {
            logger.error(e.getMessage(),e);
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException e1) {
                    logger.error(e1.getMessage(),e1);
                }
            }
        }
        return sbf.toString();
    }
}
