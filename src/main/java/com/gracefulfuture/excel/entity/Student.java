package com.gracefulfuture.excel.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @author chenkun
 * @version 1.0
 * @description 学生实体类
 * @create 2021/7/23 10:14
 */
@Data
public class Student {
    /**
     * 学号
     */
    @ExcelProperty(value = "学号",index = 0)
    private String number;
    /**
     * 姓名
     */
    @ExcelProperty(value = "姓名",index = 1)
    private String fullName;
    /**
     * 年龄
     */
    @ExcelProperty(value = "年龄",index = 2)
    private Integer age;
    /**
     * 班号
     */
    @ExcelProperty(value = "班号",index = 3)
    private String classNo;
    /**
     * 地址
     */
    @ExcelProperty(value = "住址",index = 4)
    private String address;
}
