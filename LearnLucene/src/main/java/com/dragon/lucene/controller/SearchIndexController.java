package com.dragon.lucene.controller;

import com.dragon.common.Constants;
import com.dragon.common.SearchParam;
import com.dragon.lucene.utils.LuceneUtils;
import org.apache.commons.collections.CollectionUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 11:20
 * @description：查询索引库
 * @modified By：
 * @version: $
 */
public class SearchIndexController {
    private static final Logger logger = LoggerFactory.getLogger(SearchIndexController.class);

    public static void main(String[] args) {
        List<Map<String, Object>> result = getSearchResult("中国");
        if(CollectionUtils.isNotEmpty(result)){
            for (Map<String, Object> map : result) {
                logger.info("查询结果："+ map);
            }
        }
    }

    private static List<Map<String, Object>> getSearchResult(String searchText) {
        Map<String, Object> resultMap = new HashMap<String,Object>();
        String filePath = LuceneUtils.getFilePath(Constants.TEST_CACHE_FILENAME);

        String[] fields = null;
        // fields = new String[]{"companyName"};  // 搜索单个字段
        fields = new String[]{"companyName","companyAddress"};  // 搜索多个字段

        List<String> backFieldList = new ArrayList<>();// 返回查询字段
        backFieldList.add("id");
        backFieldList.add("companyName");
        backFieldList.add("companyAddress");

        SearchParam sp = new SearchParam();
        sp.getSp().put("fields", fields);
        sp.getSp().put("filePath", filePath);
        sp.getSp().put("backFieldList", backFieldList);
        sp.getSp().put("proName", searchText);
        sp.setPageNo(1);

        List<Map<String,Object>> returnList = LuceneUtils.getLuceneSearchInfo(sp);
        if(CollectionUtils.isNotEmpty(returnList)){
            for (Map<String, Object> returnMap : returnList) {
                String id = LuceneUtils.getStringToMap(returnMap,"id");
                String companyName = LuceneUtils.getStringToMap(returnMap,"companyName");
                String companyAddress = LuceneUtils.getStringToMap(returnMap,"companyAddress");

                resultMap.put("id", id);
                resultMap.put("companyName", companyName);
                resultMap.put("companyAddress", companyAddress);
            }

        }
        return returnList;
    }
}
