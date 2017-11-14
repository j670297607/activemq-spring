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

}