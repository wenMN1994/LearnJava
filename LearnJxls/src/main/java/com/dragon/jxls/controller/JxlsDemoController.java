package com.dragon.jxls.controller;

import com.dragon.common.entity.Course;
import com.dragon.common.entity.Student;
import com.dragon.jxls.utils.JxlsUtils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 14:13
 * @description：
 * @modified By：
 * @version: $
 */
public class JxlsDemoController {
    public static void main( String[] args ) throws Exception {

        List<Student> students = new ArrayList<>();

        Student student = new Student();
        student.setName("张三");
        student.setGender("男");
        student.setGradeClass("高三（1）班");
        List<Course> courses = new ArrayList<>();
        Course course = new Course();
        course.setCourseName("语文");
        course.setCourseScore("98");
        courses.add(course);
        course = new Course();
        course.setCourseName("数学");
        course.setCourseScore("105");
        courses.add(course);
        course = new Course();
        course.setCourseName("物理");
        course.setCourseScore("80");
        courses.add(course);

        student.setCourses(courses);
        students.add(student);


        student = new Student();
        student.setName("王丽丽");
        student.setGender("女");
        student.setGradeClass("高三（2）班");
        courses = new ArrayList<>();
        course = new Course();
        course.setCourseName("语文");
        course.setCourseScore("102");
        courses.add(course);
        course = new Course();
        course.setCourseName("数学");
        course.setCourseScore("110");
        courses.add(course);
        student.setCourses(courses);
        students.add(student);

        student = new Student();
        student.setName("李梅");
        student.setGender("女");
        student.setGradeClass("高三（3）班");
        courses = new ArrayList<>();
        course = new Course();
        course.setCourseName("语文");
        course.setCourseScore("110");
        courses.add(course);
        course = new Course();
        course.setCourseName("数学");
        course.setCourseScore("100");
        courses.add(course);
        course = new Course();
        course.setCourseName("物理");
        course.setCourseScore("85");
        courses.add(course);
        student.setCourses(courses);
        students.add(student);

        //模板里展示的数据
        Map<String, Object> data = new HashMap<>();
        data.put("students", students);

        // 模板路径和输出流
        String templatePath = "D:\\Project\\idea-workspace\\LearnJava\\LearnJxls\\src\\main\\resources\\StudentTemplate.xlsx";
        String createFilePath = "D:\\student.xlsx";

        //调用封装的工具类，传入模板路径，文档输出路径，和装有数据的Map,按照模板导出
        JxlsUtils.exportExcel(templatePath, createFilePath, data);
    }
}
