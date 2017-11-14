package com.annotation;


import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;
import com.google.common.collect.Maps;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;
import com.annotation.Excel.Struct;
import com.annotation.Excel.TitleType;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * Created by Joanna on 2017/11/13.
 */
public class ExcelUtil<T> {
    private final static String FIELD_CACHE_PREFIX = "f_";
    private static Cache<String, Map<String, String>> fieldCache = CacheBuilder.newBuilder().expireAfterAccess(1L, TimeUnit.HOURS).build();

    public static <ModelType> Excel getExcel(Class<ModelType> clazz) {
        return clazz.getAnnotation(Excel.class);
    }

    /**
     * 获取excel采用何种方式配置
     * @param clazz
     * @return
     */
    public static <ModelType> Struct getExcelStruct(Class<ModelType> clazz) {
        Struct struct = Struct.ANNOTATION;
        Excel excel = getExcel(clazz);
        if (null != excel) {
            struct = excel.struct();
        }
        return struct;
    }

    public static <ModelType> Map<String, String> getExcelCells(final Class<ModelType> clazz) {
        try {
            return fieldCache.get(FIELD_CACHE_PREFIX + clazz.getSimpleName(), new Callable<Map<String, String>>() {
                @Override
                public Map<String, String> call() throws Exception {
                    Struct struct = getExcelStruct(clazz);
                    Configuration configuration = null;
                    if (Struct.ANNOTATION == struct) {
                        configuration = new AnnotationConfiguration();
                    } else {
                        configuration = new XmlConfiguration();
                    }
                    return configuration.parse(clazz);
                }
            });
        } catch (ExecutionException e) {
        }
        return Maps.newHashMap();
    }

    static XSSFReader open(InputStream in) throws Exception {
        OPCPackage pkg = OPCPackage.open(in);
        return new XSSFReader(pkg);
    }
    static XMLReader getXMLReader() throws SAXException {
        return XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
    }
    static <ModelType> SheetHandler<ModelType> getHandler(Class<ModelType> clazz, SharedStringsTable sst)
            throws SAXException {
        return new SheetHandler<ModelType>(clazz, sst);
    }

    public static <ModelType> List<ModelType> importSheet(InputStream in, Class<ModelType> clazz) throws Exception {
        XSSFReader r = open(in);
        SharedStringsTable sst = r.getSharedStringsTable();
        XMLReader parser = getXMLReader();
        SheetHandler<ModelType> handler = getHandler(clazz, sst);
        parser.setContentHandler(handler);
        InputStream sheet1 = r.getSheet("rId1");
        InputSource sheetSource = new InputSource(sheet1);
        parser.parse(sheetSource);
        sheet1.close();
        return handler.getData();
    }

    public static <ModelType> TitleType getExcelTitleType(Class<ModelType> clazz) {
        TitleType titleType = TitleType.SIMPLE;
        Excel excel = getExcel(clazz);
        if (null != excel) {
            titleType = excel.titleType();
        }
        return titleType;
    }

    public static <ModelType> int getExcelDataStart(Class<ModelType> clazz) {
        int start = 2;
        Excel excel = getExcel(clazz);
        if (null != excel) {
            start = excel.start();
        }
        return start;
    }






//    List<T> importExcel(String sheetName, InputStream input) {
//        int maxCol = 0;
//        List<T> list = new ArrayList<T>();
//        try {
//            Workbook workbook = WorkbookFactory.create(input);
//            Sheet sheet = workbook.getSheet(sheetName);
//            // 如果指定sheet名,则取指定sheet中的内容.
//            if (!sheetName.trim().equals("")) {
//                sheet = workbook.getSheet(sheetName);
//            }
//            // 如果传入的sheet名不存在则默认指向第1个sheet.
//            if (sheet == null) {
//                sheet = workbook.getSheetAt(0);
//            }
//            int rows = sheet.getPhysicalNumberOfRows();
//            // 有数据时才处理
//            if (rows > 0) {
//                List<Field> allFields = getMappedFiled(clazz, null);
//                // 定义一个map用于存放列的序号和field.
//                Map<Integer, Field> fieldsMap = new HashMap<Integer, Field>();
//                // 第一行为表头
//                Row rowHead = sheet.getRow(0);
//                Map<String, Integer> cellMap = new HashMap<String, Integer>();
//                int cellNum = rowHead.getPhysicalNumberOfCells();
//                for (int i = 0; i < cellNum; i++) {
//                    cellMap.put(rowHead.getCell(i).getStringCellValue().toLowerCase(), i);
//                }
//                for (Field field : allFields) {
//                    // 将有注解的field存放到map中.
//                    if (field.isAnnotationPresent(Excel.class)) {
//                        Excel attr = field.getAnnotation(Excel.class);
//                        // 根据Name来获取相应的failed
//                        int col = cellMap.get(attr.name().toLowerCase());
//                        field.setAccessible(true);
//                        fieldsMap.put(col, field);
//                    }
//                }
//                // 从第2行开始取数据
//                for (int i = 1; i < rows; i++) {
//                    Row row = sheet.getRow(i);
//                    T entity = null;
//                    for (int j = 0; j < cellNum; j++) {
//                        Cell cell = row.getCell(j);
//                        if (cell == null) {
//                            continue;
//                        }
//                        int cellType = cell.getCellType();
//                        String c = "";
//                        if (cellType == HSSFCell.CELL_TYPE_NUMERIC) {
//                            DecimalFormat df = new DecimalFormat("0");
//                            c = df.format(cell.getNumericCellValue());
//                        } else if (cellType == HSSFCell.CELL_TYPE_BOOLEAN) {
//                            c = String.valueOf(cell.getBooleanCellValue());
//                        } else {
//                            c = cell.getStringCellValue();
//                        }
//                        if (c == null || c.equals("")) {
//                            continue;
//                        }
//                        entity = (entity == null ? clazz.newInstance() : entity);
//                        // 从map中得到对应列的field.
//                        Field field = fieldsMap.get(j);
//                        if (field == null) {
//                            continue;
//                        }
//                        // 取得类型,并根据对象类型设置值.
//                        Class<?> fieldType = field.getType();
//                        if (String.class == fieldType) {
//                            field.set(entity, String.valueOf(c));
//                        } else if ((Integer.TYPE == fieldType)
//                                || (Integer.class == fieldType)) {
//                            field.set(entity, Integer.valueOf(c));
//                        } else if ((Long.TYPE == fieldType)
//                                || (Long.class == fieldType)) {
//                            field.set(entity, Long.valueOf(c));
//                        } else if ((Float.TYPE == fieldType)
//                                || (Float.class == fieldType)) {
//                            field.set(entity, Float.valueOf(c));
//                        } else if ((Short.TYPE == fieldType)
//                                || (Short.class == fieldType)) {
//                            field.set(entity, Short.valueOf(c));
//                        } else if ((Double.TYPE == fieldType)
//                                || (Double.class == fieldType)) {
//                            field.set(entity, Double.valueOf(c));
//                        } else if (Character.TYPE == fieldType) {
//                            if (c.length() > 0) {
//                                field.set(entity, c.charAt(0));
//                            }
//                        }
//                    }
//                    if (entity != null) {
//                        list.add(entity);
//                    }
//                }
//            }
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        return list;
//    }
//
//    /**
//     * 得到实体类所有通过注解映射了数据表的字段
//     *
//     * @param clazz
//     * @param fields
//     * @return
//     */
//    private List<Field> getMappedFiled(Class clazz, List<Field> fields) {
//        if (fields == null) {
//            fields = new ArrayList<Field>();
//        }
//        // 得到所有定义字段
//        Field[] allFields = clazz.getDeclaredFields();
//        // 得到所有field并存放到一个list中.
//        for (Field field : allFields) {
//            if (field.isAnnotationPresent(Excel.class)) {
//                fields.add(field);
//            }
//        }
//        if (clazz.getSuperclass() != null && !clazz.getSuperclass().equals(Object.class)) {
//            getMappedFiled(clazz.getSuperclass(), fields);
//        }
//        return fields;
//    }

}