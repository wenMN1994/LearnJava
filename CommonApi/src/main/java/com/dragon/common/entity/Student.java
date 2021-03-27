package com.dragon.common.entity;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 14:14
 * @description：
 * @modified By：
 * @version: $
 */
public class Student implements Serializable {

    private static final long serialVersionUID = 1L;

    private String name;

    private String gender;

    private String gradeClass;

    private List<Course> courses = new ArrayList<>();

    @Override
    public String toString() {
        return "Student{" + "name='" + name + '\'' + ", gender='" + gender + '\'' + ", gradeClass='" + gradeClass + '\'' + ", courses=" + courses + '}';
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public String getGradeClass() {
        return gradeClass;
    }

    public void setGradeClass(String gradeClass) {
        this.gradeClass = gradeClass;
    }

    public List<Course> getCourses() {
        return courses;
    }

    public void setCourses(List<Course> courses) {
        this.courses = courses;
    }
}

