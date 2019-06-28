package com.mbc.common.utils.excel.model;

/**
 * Created by 符柱成 on 2017/8/24.
 */
public class StudentModelExcelTest {
    private String name;
    private int id;
    private String sex;

    public StudentModelExcelTest(int id, String name, String sex) {
        this.id = id;
        this.name = name;
        this.sex = sex;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
