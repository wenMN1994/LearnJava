package com.dragon.lucene.utils;

import com.dragon.common.Constants;
import com.dragon.common.SearchParam;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.lucene.analysis.Analyzer;
import org.apache.lucene.analysis.standard.StandardAnalyzer;
import org.apache.lucene.document.Document;
import org.apache.lucene.index.DirectoryReader;
import org.apache.lucene.index.IndexReader;
import org.apache.lucene.queryparser.classic.MultiFieldQueryParser;
import org.apache.lucene.queryparser.classic.QueryParser.Operator;
import org.apache.lucene.search.BooleanClause.Occur;
import org.apache.lucene.search.*;
import org.apache.lucene.store.Directory;
import org.apache.lucene.store.FSDirectory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 11:16
 * @description：LuceneUtils工具类
 * @modified By：
 * @version: $
 */
public class LuceneUtils {

    private static final Logger logger = LoggerFactory.getLogger(LuceneUtils.class);

    /**
     * 每个结果集单独创建一个路径
     * @param rootFilePath
     * @return
     */
    public static String getFilePath(String rootFilePath){
        // Windows系统下默认为D盘根目录下D:\opt\files\dragon
        String img_Url = "/opt/files/dragon";
        String FILE_NAME = Constants.TEST_CACHE_FILENAME_PATH;
        String filePath = img_Url + "/" + FILE_NAME + "/" + rootFilePath + "/";
        File targetFile = new File(filePath);
        if (!targetFile.exists()) {
            targetFile.mkdirs();
        }
        return filePath;
    }


    /**
     * 查询内容的公共方法
     * 说明：
     * pageNo：分页页码（必传）
     * limit：分页大小，查询全部时候传-1（必传）
     * proName：搜索内容参数
     * filePath：索引文件路径（必传）
     * fields：搜索字段（必传）
     * backFieldList：查询返回数据的字段集合List（必传）
     * queryList: 组合查询逻辑（非必传）
     * @param sp
     * @return
     */
    public static List<Map<String,Object>> getLuceneSearchInfo(SearchParam sp){
        // 查询信息（公共方法）
        List<Map<String,Object>> returnList = new ArrayList<>();
        int pageNo = sp.getPageNo();
        int limit = sp.getLimit();
        int total = sp.getSp().get("total") == null ? 0 : (int) sp.getSp().get("total");
        // 搜索内容（非必传）
        String proName = sp.getSp().get("proName") == null ? "" : sp.getSp().get("proName")+"";
        // 搜索条件（非必传）
        String isOr = sp.getSp().get("isOr") == null ? "" : sp.getSp().get("isOr")+"";
        // 索引文件路径（必传）
        String filePath = sp.getSp().get("filePath") == null ? "" : sp.getSp().get("filePath")+"";
        // 搜索字段数组（必传）
        String[] fields = sp.getSp().get("fields") == null ? new String[]{} : (String[]) sp.getSp().get("fields");
        // 查询索引文档的数据字段（必传）
        @SuppressWarnings("unchecked")
        List<String> backFieldList = sp.getSp().get("backFieldList") == null ? new ArrayList<>() : (ArrayList<String>) sp.getSp().get("backFieldList");
        // (或)query 对象（非必传）
        @SuppressWarnings("unchecked")
        List<Query> queryList = sp.getSp().get("queryList") == null ? new ArrayList<>() : (ArrayList<Query>) sp.getSp().get("queryList");
        // (并且)query 对象（非必传）
        @SuppressWarnings("unchecked")
        List<Query> queryMustList = sp.getSp().get("queryMustList") == null ? new ArrayList<>() : (ArrayList<Query>) sp.getSp().get("queryMustList");
        // (并且)query 对象（非必传）
        @SuppressWarnings("unchecked")
        List<Query> queryMustNotList = sp.getSp().get("queryMustNotList") == null ? new ArrayList<>() : (ArrayList<Query>) sp.getSp().get("queryMustNotList");

        Sort sort = sp.getSp().get("sort") == null ? null : (Sort) sp.getSp().get("sort");

        if(StringUtils.isBlank(filePath)){
            throw new RuntimeException("filePath：索引文档路径为空");
        }
        if(fields.length == 0){
            throw new RuntimeException("fields：搜索字段数组为空");
        }

        if(StringUtils.isBlank(proName)){
            // 为空的时候，检索出全部内容
            proName = "*:*";
        }
        try {
            Path indexPath = Paths.get(filePath);
            Directory dir = FSDirectory.open(indexPath);
            IndexReader reader = DirectoryReader.open(dir);
            IndexSearcher searcher = new IndexSearcher(reader);
            // 当为true时，分词器进行智能切分
            Analyzer analyzer = new StandardAnalyzer();
            // 多域搜索
            MultiFieldQueryParser parser = new MultiFieldQueryParser(fields, analyzer);
            if(StringUtils.isBlank(isOr)) {
                // 关键字同时成立使用 AND, 默认是 OR（搜）
                parser.setDefaultOperator(Operator.AND);
            }else {
                //默认是 OR
                parser.setDefaultOperator(Operator.OR);
            }
            BooleanQuery.Builder builder = new BooleanQuery.Builder();

            // 查询关键词
            Query query = parser.parse(proName);
            builder.add(query, Occur.MUST);

            if(CollectionUtils.isNotEmpty(queryMustList)){
                for (Query queryTemp : queryMustList) {
                    if(queryTemp == null){
                        continue;
                    }
                    // 关联其它域 (Occur.MUST:并且；Occur.SHOULD：或者；Occur.MUST_NOT：非)
                    builder.add(queryTemp, Occur.MUST);
                }
            }
            if(CollectionUtils.isNotEmpty(queryMustNotList)){
                for (Query queryTemp : queryMustNotList) {
                    if(queryTemp == null){
                        continue;
                    }
                    // 关联其它域 (Occur.MUST:并且；Occur.SHOULD：或者；Occur.MUST_NOT：非)
                    builder.add(queryTemp, Occur.MUST_NOT);
                }
            }

            TopDocs tds = null;
            if(sort == null){
                // 返回查询的数量
                tds = searcher.search(builder.build(), pageNo * limit);
            }else{
                // （排序）返回查询的数量
                tds = searcher.search(builder.build(), pageNo * limit, sort);
            }

            ScoreDoc[] docs = tds.scoreDocs;
            int totalHits = (int) tds.totalHits;
            total = totalHits;
            // 返回总条数
            sp.getSp().put("total", total);

            //判断是否超过总记录数
            int count = 0;
            // 返回全部数据
            if(limit == -1){
                count = totalHits;
            }else{
                if (totalHits <= pageNo * limit) {
                    count = totalHits;
                }  else {
                    count = pageNo*limit;
                }
            }

            // 取出当前页的数据
            for (int i = ((pageNo-1) * limit); i < count; i++ ){
                ScoreDoc doc = docs[i];
                // 文档id
                int docId = doc.doc;
                // 文档评分
                //float score = doc.score;
                Document document = searcher.doc(docId);
                Map<String,Object> qryMap = new HashMap<>();
                for(String qryField : backFieldList){
                    Object qryValue = document.get(qryField) == null ? "" : document.get(qryField);
                    qryMap.put(qryField, qryValue);
                }
                returnList.add(qryMap);
            }
            dir.close();
            reader.close();
        } catch (Exception e) {
            logger.error(e.getMessage(),e);
        }
        return returnList;
    }

    /**
     * Map： 对null字符处理
     * @return
     */
    public static String getStringToMap(Map<String, Object> map,String key){
        String returnStr = "";
        if(map == null || StringUtils.isBlank(key)){
            return returnStr;
        }
        returnStr = map.get(key) == null ? "" : map.get(key) + "";
        return returnStr;
    }

}
