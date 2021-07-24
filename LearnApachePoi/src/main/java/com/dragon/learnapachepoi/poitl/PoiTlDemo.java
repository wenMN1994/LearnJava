package com.dragon.learnapachepoi.poitl;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.DocxRenderData;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/6/27 21:50
 * @description：API:http://deepoove.com/poi-tl/
 * @modified By：
 * @version: $
 */
public class PoiTlDemo {

    public static void main(String[] args) throws IOException {
        XWPFTemplate template = XWPFTemplate.compile("D:\\Project\\idea-workspace\\LearnJava\\LearnApachePoi\\src\\main\\resources\\wordTemplate\\NewTemplate.docx").render(
                new HashMap<String, Object>(){{
//                    put("title", "Hi, poi-tl Word模板引擎");
                    put("product3_1_1", new DocxRenderData(new File("D:\\Data\\39.word\\word3_1_1.docx")));
                    put("product3_1_2", new DocxRenderData(new File("D:\\Data\\39.word\\word3_1_2.docx")));
                    put("product3_2_1", new DocxRenderData(new File("D:\\Data\\39.word\\word3_2_1.docx")));
                    put("product3_2_2", new DocxRenderData(new File("D:\\Data\\39.word\\word3_2_2.docx")));
                }});

        FileOutputStream out = new FileOutputStream("D:\\poi_tl_output.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
