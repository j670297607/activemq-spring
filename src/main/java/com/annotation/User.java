package com.annotation;

/**
 * Created by user on 2017/11/13.
 */
@Excel(start=2)
public class User {

    @ExcelCell("filed1")
    private String name;
    @ExcelCell("filed2")
    private String nameEn;
    @ExcelCell("filed3")
    private Integer age;
    @ExcelCell("filed4")
    private String six;
    @ExcelCell("filed5")
    private String weight;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getNameEn() {
        return nameEn;
    }

    public void setNameEn(String nameEn) {
        this.nameEn = nameEn;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getSix() {
        return six;
    }

    public void setSix(String six) {
        this.six = six;
    }

    public String getWeight() {
        return weight;
    }

    public void setWeight(String weight) {
        this.weight = weight;
    }

    @Override
    public String toString() {
        return "User{" +
                "name='" + name + '\'' +
                ", nameEn='" + nameEn + '\'' +
                ", age=" + age +
                ", six='" + six + '\'' +
                ", weight='" + weight + '\'' +
                '}';
    }
}
