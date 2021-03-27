package com.dragon.common.entity;

import java.io.Serializable;

/**
 * @author：Dragon Wen
 * @email：18475536452@163.com
 * @date：Created in 2021/3/21 14:11
 * @description：
 * @modified By：
 * @version: $
 */
public class Course implements Serializable {

    private static final long serialVersionUID = 1L;

    private String courseName;

    private String courseScore;

    @Override
    public String toString() {
        return "Course{" + "courseName='" + courseName + '\'' + ", courseScore='" + courseScore + '\'' + '}';
    }

    public String getCourseName() {
        return courseName;
    }

    public void setCourseName(String courseName) {
        this.courseName = courseName;
    }

    public String getCourseScore() {
        return courseScore;
    }

    public void setCourseScore(String courseScore) {
        this.courseScore = courseScore;
    }
}

