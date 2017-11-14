package com.annotation;

import com.google.common.collect.Lists;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.lang.reflect.ParameterizedType;
import java.util.List;
import java.util.Map;

/**
 * Created by user on 2017/11/13.
 */
public class ImportTest {
    public static void main(String[] args) {
        new ImportTest().importExcel();
    }

    public void importExcel() {
        try {
            List<User> list = ExcelUtil.importSheet(new FileInputStream("D://data.xlsx"), User.class);
            for (User u : list) {
                System.out.println(u.toString());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
