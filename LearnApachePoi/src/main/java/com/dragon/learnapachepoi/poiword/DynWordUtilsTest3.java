package com.dragon.learnapachepoi.poiword;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author DragonWen
 */
public class DynWordUtilsTest3 {

    public static void main(String[] args) {
        // 模板全的路径
        String templatePath = "D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\申请资料.docx";
        // 输出位置
        String outPath = "D:\\申请资料.doc";

        Map<String, Object> paramMap = new HashMap<>(16);

        // 申请资料table
        List<List<String>> tbAddRowCopyRow = new ArrayList<>();
        // table 数据组装
        for (int i = 1; i <= 2; i++) {
            List<String> tempCopyRowList = new ArrayList<>();
            tempCopyRowList.add("测试"+i);
            tempCopyRowList.add("√");
            tempCopyRowList.add("√");
            tempCopyRowList.add("√");
            tempCopyRowList.add("");
            tempCopyRowList.add("以实际时间为准");
            tbAddRowCopyRow.add(tempCopyRowList);
        }

        paramMap.put("tbInsertRow:6,7,copyRow", tbAddRowCopyRow);
        DynWordUtils.process(paramMap, templatePath, outPath);
    }
}
