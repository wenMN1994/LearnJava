package com.dragon.learnapachepoi.controller;

import com.dragon.common.Result;
import com.dragon.common.entity.UserEntity;
import com.dragon.learnapachepoi.uitls.ExcelData;
import com.dragon.learnapachepoi.uitls.ExportExcelUtils;
import com.dragon.learnapachepoi.uitls.ParseExcelUtils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 10:45
 * @description：
 * @modified By：
 * @version: $
 */
@Controller
@RequestMapping("/excel")
public class ExcelController {

    /**
     * 导出单个Sheet的Excel
     */
    @RequestMapping("/exportExcelSingleSheet")
    @ResponseBody
    public String exportExcelSingleSheet(HttpServletResponse res){
        try {
            List<List<Object>> dataList = new ArrayList<List<Object>>();
            ExcelData excelData = new ExcelData();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddhhmmss");
            String str=sdf.format(new Date());

            List<UserEntity> userEntityList = new ArrayList<>();
            for (int i = 0; i < 5; i++) {
                UserEntity userEntity = new UserEntity();
                userEntity.setName("张三"+i);
                userEntity.setAge("2"+i+"");
                if(i%2 == 0){
                    userEntity.setSex("男");
                }else {
                    userEntity.setSex("女");
                }
                userEntityList.add(userEntity);
            }

            List<String> titles=new ArrayList<String>();
            titles.add("");
            titles.add("姓名");
            titles.add("年龄");
            titles.add("性别");
            int i=1;

            for (UserEntity userEntity : userEntityList) {
                List<Object> data = new ArrayList<Object>();
                data.add(i+"");
                data.add(userEntity.getName() == null ? "-" : userEntity.getName());
                data.add(userEntity.getAge() == null ? "-" : userEntity.getAge());
                data.add(userEntity.getSex() == null ? "-" : userEntity.getSex());
                i++;
                dataList.add(data);
            }

            excelData.setTitles(titles);
            excelData.setRows(dataList);
            excelData.setName("用户列表");
            ExportExcelUtils.exportExcel(res,"用户列表"+str+ ".xlsx", excelData);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return "success";
    }

    /**
     * 导出多个Sheet的Excel
     */
    @RequestMapping("/exportExcelMultiSheet")
    @ResponseBody
    public String exportExcelMultiSheet(HttpServletResponse res){
        try {
            List<ExcelData> dataList = new ArrayList<>();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddhhmmss");
            String str=sdf.format(new Date());

            List<UserEntity> userEntityList = new ArrayList<>();
            for (int i = 0; i < 5; i++) {
                UserEntity userEntity = new UserEntity();
                userEntity.setName("张三"+i);
                userEntity.setAge("2"+i+"");
                if(i%2 == 0){
                    userEntity.setSex("男");
                }else {
                    userEntity.setSex("女");
                }
                userEntityList.add(userEntity);
            }

            List<List<Object>> oneDataList = new ArrayList<List<Object>>();
            ExcelData oneExcelData = new ExcelData();
            List<String> oneTitles = new ArrayList<String>();
            oneTitles.add("");
            oneTitles.add("姓名");
            oneTitles.add("年龄");
            oneTitles.add("性别");
            int i=1;

            for (UserEntity userEntity : userEntityList) {
                List<Object> data = new ArrayList<Object>();
                data.add(i+"");
                data.add(userEntity.getName() == null ? "-" : userEntity.getName());
                data.add(userEntity.getAge() == null ? "-" : userEntity.getAge());
                data.add(userEntity.getSex() == null ? "-" : userEntity.getSex());
                i++;
                oneDataList.add(data);
            }

            oneExcelData.setTitles(oneTitles);
            oneExcelData.setRows(oneDataList);
            oneExcelData.setName("用户列表1");
            dataList.add(oneExcelData);

            List<List<Object>> twoDataList = new ArrayList<List<Object>>();
            ExcelData twoExcelData = new ExcelData();
            List<String> twoTitles = new ArrayList<String>();
            twoTitles.add("");
            twoTitles.add("姓名");
            twoTitles.add("年龄");
            twoTitles.add("性别");
            int j=1;

            for (UserEntity userEntity : userEntityList) {
                List<Object> data = new ArrayList<Object>();
                data.add(j+"");
                data.add(userEntity.getName() == null ? "-" : userEntity.getName());
                data.add(userEntity.getAge() == null ? "-" : userEntity.getAge());
                data.add(userEntity.getSex() == null ? "-" : userEntity.getSex());
                j++;
                twoDataList.add(data);
            }

            twoExcelData.setTitles(twoTitles);
            twoExcelData.setRows(twoDataList);
            twoExcelData.setName("用户列表2");
            dataList.add(twoExcelData);


            ExportExcelUtils.exportExcelMultiSheet(res,"用户列表"+str+ ".xlsx", dataList);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return "success";
    }

    /**
     * excel导入
     * @param excelFile
     * @return
     */
    @RequestMapping("/excelImport")
    @ResponseBody
    public Result excelImport(@RequestParam("excelFile") MultipartFile excelFile) {
        Result result = Result.SUCCESS();
        InputStream fis = null;
        try {
            if(excelFile==null){
                throw new Exception("请选择excel文件");
            }

            // 读取input流
            fis = excelFile.getInputStream();
            //读取文件名称
            String fileName = excelFile.getOriginalFilename();
            // 获取文件后缀
            String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
            List<List<String>> dateList=new ArrayList<List<String>>();
            if (suffix.equals("xlsx")) {
                //07版
                dateList = ParseExcelUtils.parseExcelxlsx(fis);
            }else if (suffix.equals("xls")) {
                //03版
                dateList = ParseExcelUtils.parseExcelxls(fis);
            }else {
                throw new Exception("请选择excel文件");
            }
            if(dateList==null || dateList.size()<=0){
                throw new RuntimeException("导入数据为空。");
            }

            List<Map<String, Object>> resultList = new ArrayList<>();
            int i=1;
            for(List<String> date : dateList){
                String number = date.get(0)==null?"":date.get(0) +"";
                String name = date.get(1)==null?"":date.get(1) +"";
                String province = date.get(2)==null?"":date.get(2) +"";
                String city = date.get(3)==null?"":date.get(3) +"";
                String salesVolume = date.get(4)==null?"":date.get(4) +"";
                String remark = date.get(5)==null?"":date.get(5) +"";
                if(StringUtils.isBlank(number)) {
                    throw new RuntimeException("第"+i+"行编号不能为空！");
                }
                if(StringUtils.isBlank(name)) {
                    throw new RuntimeException("第"+i+"行名称不能为空！");
                }
                if(StringUtils.isBlank(province)) {
                    throw new RuntimeException("第"+i+"行省份不能为空！");
                }
                if(StringUtils.isBlank(city)) {
                    throw new RuntimeException("第"+i+"行城市不能为空！");
                }
                if(StringUtils.isBlank(salesVolume)) {
                    throw new RuntimeException("第"+i+"行前一日销量不能为空！");
                }

                Map<String, Object> itemMap = new HashMap<>();
                itemMap.put("number",number);
                itemMap.put("name",name);
                itemMap.put("province",province);
                itemMap.put("city",city);
                itemMap.put("salesVolume",salesVolume);
                itemMap.put("remark",remark);
                resultList.add(itemMap);
                i++;
            }
            result.data = resultList;
            return result;

        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

}
