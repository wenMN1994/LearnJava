package com.dragon.learnapachepoi.poiword;

import java.util.*;


/**
 * @author DragonWen
 */
public class DynWordUtilsTest2 {
	
	/**
     * 说明 普通占位符位${field}格式
     * 表格中的占位符为${tbAddRow:tb1}  tb1为唯一标识符
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) {

        // 模板全的路径
//        String templatePath = "D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\新建 Microsoft Word 文档.docx";
        String templatePath = "D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\word00.docx";

    	// 输出位置
        String outPath = "D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\NewTemplate.docx";
        
        Map<String, Object> paramMap = new HashMap<>(16);
        // 普通的占位符示例 参数数据结构 {str,str}
        paramMap.put("title", "德玛西亚");
        paramMap.put("startYear", "2010");
        paramMap.put("endYear", "2020");
        paramMap.put("currentYear", "2019");
        paramMap.put("currentMonth", "10");
        paramMap.put("currentDate", "26");
        paramMap.put("name", "黑色玫瑰");

        // 段落中的动态段示例 [str], 支持动态行中添加图片
        List<Object> list1 = new ArrayList<>(Arrays.asList("2、list1_11111", "3、list1_2222", "${image:image0}"));
        ImageEntity imgEntity = new ImageEntity();
        imgEntity.setHeight(200);
        imgEntity.setWidth(300);
        imgEntity.setUrl("D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\image1.jpg");
        imgEntity.setTypeId(ImageUtils.ImageType.JPG);
        // paramMap.put("image:image0", imgEntity);
        // paramMap.put("list1", list1);

        List<String> list2 = new ArrayList<>(Arrays.asList("2、list2_11111", "3、list2_2222"));
        // paramMap.put("list2", list2);

        // 表格中的参数示例 参数数据结构 [[str]]
        List<List<String>> tbRow1 = new ArrayList<>();
        List<String> tbRow1_row1 = new ArrayList<>(Arrays.asList("渠道A", "www.qudaoa.com", "0755-2654235"));
        List<String> tbRow1_row2 = new ArrayList<>(Arrays.asList("渠道B", "www.qudaob.com", "0755-2654235"));
        List<String> tbRow1_row3 = new ArrayList<>(Arrays.asList("渠道C", "www.qudaoc.com", "0755-2654235"));
        tbRow1.add(tbRow1_row1);
        tbRow1.add(tbRow1_row2);
        tbRow1.add(tbRow1_row3);
        paramMap.put(PoiWordUtils.addRowText + "tb1", tbRow1);

        List<List<String>> tbRow2 = new ArrayList<>();
        List<String> tbRow2_row1 = new ArrayList<>(Arrays.asList("指标c", "指标c的意见"));
        List<String> tbRow2_row2 = new ArrayList<>(Arrays.asList("指标d", "指标d的意见"));
        tbRow2.add(tbRow2_row1);
        tbRow2.add(tbRow2_row2);
        paramMap.put(PoiWordUtils.addRowText + "tb2", tbRow2);

        List<Map<String,List<List<String>>>> dataAddList = new ArrayList<>();
        // List<List<String>> tbRow3 = new ArrayList<>();

        List<String> tbRow3_row1 = new ArrayList<>(Arrays.asList("认购费率(010688)", "0万元 ≤ 金额 < 100万元","1%"));
        List<String> tbRow3_row2 = new ArrayList<>(Arrays.asList("认购费率(010688)", "100万元 ≤ 金额 < 300万元","1.8%"));
        List<String> tbRow3_row3 = new ArrayList<>(Arrays.asList("认购费率(010688)", "300万元 ≤ 金额 < 500万元","1.6%"));
        List<String> tbRow3_row4 = new ArrayList<>(Arrays.asList("认购费率(010688)", "500万元 ≤ 金额","每笔1000元"));

        Map<String,List<List<String>>> map1 = new HashMap<>();
        List<List<String>> tbRow31 = new ArrayList<>();
        tbRow31.add(tbRow3_row1);
        tbRow31.add(tbRow3_row2);
        tbRow31.add(tbRow3_row3);
        tbRow31.add(tbRow3_row4);
        map1.put("tbAddRowMergeRow1:0,20,23", tbRow31);
        dataAddList.add(map1);

        List<String> tbRow3_row5 = new ArrayList<>(Arrays.asList("申购费率(010688)", "0万元 ≤ 金额 < 100万元","0.2%"));
        List<String> tbRow3_row6 = new ArrayList<>(Arrays.asList("申购费率(010688)", "100万元 ≤ 金额 < 300万元","0.1%"));
        List<String> tbRow3_row7 = new ArrayList<>(Arrays.asList("申购费率(010688)", "300万元 ≤ 金额 < 500万元","0.8%"));
        List<String> tbRow3_row8 = new ArrayList<>(Arrays.asList("申购费率(010688)", "500万元 ≤ 金额","每笔1000元"));

        Map<String,List<List<String>>> map2 = new HashMap<>();
        List<List<String>> tbRow32 = new ArrayList<>();
        tbRow32.add(tbRow3_row5);
        tbRow32.add(tbRow3_row6);
        tbRow32.add(tbRow3_row7);
        tbRow32.add(tbRow3_row8);
        map2.put("tbAddRowMergeRow1:0,24,27", tbRow32);
        dataAddList.add(map2);

        List<String> tbRow3_row9 = new ArrayList<>(Arrays.asList("赎回费率(010689)", "无费用",""));
        Map<String,List<List<String>>> map3 = new HashMap<>();
        List<List<String>> tbRow33 = new ArrayList<>();
        tbRow33.add(tbRow3_row9);
        map3.put("tbAddRowMergeCol1:28,1,2", tbRow33);
        dataAddList.add(map3);

        //paramMap.put(PoiWordUtils.addRowMergeText + "tb3", tbRow3);
        paramMap.put("tbAddRowMerge:feeTable", dataAddList);

        // 支持在表格中动态添加图片
        List<List<String>> tbRow4 = new ArrayList<>();
        List<String> tbRow4_row1 = new ArrayList<>(Arrays.asList("03", "旅游用地", "18.8m2"));
        List<String> tbRow4_row2 = new ArrayList<>(Arrays.asList("04", "建筑用地"));
        List<String> tbRow4_row3 = new ArrayList<>(Arrays.asList("04", "${image:image3}"));
        tbRow4.add(tbRow4_row3);
        tbRow4.add(tbRow4_row1);
        tbRow4.add(tbRow4_row2);

        // 支持在表格中添加重复模板的行
        List<List<String>> tbRow5 = new ArrayList<>();
        List<String> tbRow5_row1 = new ArrayList<>(Arrays.asList("欢乐喜剧人"));
        List<String> tbRow5_row2 = new ArrayList<>(Arrays.asList("常远", "艾伦"));
        List<String> tbRow5_row3 = new ArrayList<>(Arrays.asList("岳云鹏", "孙越"));

        List<String> tbRow5_row4 = new ArrayList<>(Arrays.asList("诺克萨斯"));
        List<String> tbRow5_row5 = new ArrayList<>(Arrays.asList("德莱文", "诺手"));
        List<String> tbRow5_row6 = new ArrayList<>(Arrays.asList("男枪", "卡特琳娜"));

        tbRow5.add(tbRow5_row1);
        tbRow5.add(tbRow5_row2);
        tbRow5.add(tbRow5_row3);
        tbRow5.add(tbRow5_row4);
        tbRow5.add(tbRow5_row5);
        tbRow5.add(tbRow5_row6);
        paramMap.put("tbAddRowRepeat:0,2,0,1", tbRow5);

        ImageEntity imgEntity3 = new ImageEntity();
        imgEntity3.setHeight(100);
        imgEntity3.setWidth(100);
        imgEntity3.setUrl("D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\image1.jpg");
        imgEntity3.setTypeId(ImageUtils.ImageType.JPG);

        // paramMap.put(PoiWordUtils.addRowText + "tb4", tbRow4);
        // paramMap.put("image:image3", imgEntity3);

        // 图片占位符示例 ${image:imageid} 比如 ${image:image1}, ImageEntity中的值就为image:image1
        // 段落中的图片
        ImageEntity imgEntity1 = new ImageEntity();
        imgEntity1.setHeight(500);
        imgEntity1.setWidth(400);
        imgEntity1.setUrl("D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\image1.jpg");
        imgEntity1.setTypeId(ImageUtils.ImageType.JPG);
        //paramMap.put("image:image1", imgEntity1);

        // 表格中的图片
        ImageEntity imgEntity2 = new ImageEntity();
        imgEntity2.setHeight(200);
        imgEntity2.setWidth(100);
        imgEntity2.setUrl("D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\image1.jpg");
        imgEntity2.setTypeId(ImageUtils.ImageType.JPG);
        //paramMap.put("image:image2", imgEntity2);


        List<String> list3_1 = new ArrayList<>(Arrays.asList("{{+product3_1_1}}", "{{+product3_1_2}}"));
        paramMap.put("list3_1", list3_1);

        List<String> list3_2 = new ArrayList<>(Arrays.asList("{{+product3_2_1}}", "{{+product3_2_2}}"));
        paramMap.put("list3_2", list3_2);
        DynWordUtils.process(paramMap, templatePath, outPath);
    }
    
    public void testImage() {

        Map<String, Object> paramMap = new HashMap<>(16);
        String templatePaht = "D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\11.docx";
        String outPath = "D:\\3.docx";
        ImageEntity imgEntity1 = new ImageEntity();
        imgEntity1.setHeight(500);
        imgEntity1.setWidth(400);
        imgEntity1.setUrl("D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\image1.jpg");
        imgEntity1.setTypeId(ImageUtils.ImageType.JPG);

        paramMap.put("image:img1", imgEntity1);
        DynWordUtils.process(paramMap, templatePaht, outPath);
    }
	

}
