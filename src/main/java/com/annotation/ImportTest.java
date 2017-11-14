package com.annotation;

import java.io.FileInputStream;
import java.util.List;

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
