package com.annotation;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import com.annotation.Excel.TitleType;

/**
 * Created by user on 2017/11/14.
 */
public class SheetHandler<ModelType> extends DefaultHandler {

    private int start = 2;//开多少行才开始算有用数据
    private int line = 1;//当前读取第几行
    private int curCell;//标示当前位于一行的哪个单个格
    private String lastContents;//单元格的值
    private boolean nextIsString;
    private ModelType modelType;
    private Class<ModelType> clazz;
    private List<ModelType> dataList;//解析后返回的model
    private Map<Integer, String> fieldmap;
    private SharedStringsTable sst;

    private String titleString;
    private Map<String,String> titlemap;
    private TitleType titleType;

    public SheetHandler(Class<ModelType> clazz, SharedStringsTable sst) {
        this.dataList = Lists.newArrayList();
        this.fieldmap = Maps.newHashMap();
        this.clazz = clazz;
        this.sst = sst;
        this.titlemap=Maps.newHashMap();
        this.titleString="";
    }



    @Override
    public void startDocument() throws SAXException {
        super.startDocument();
        titleType=ExcelUtil.getExcelTitleType(clazz);
    }



    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        if (name.equals("row")) { // row => new line
            start = ExcelUtil.getExcelDataStart(clazz);
            if (!isTitleLine()) {
                try {
                    this.modelType = clazz.newInstance();
                    this.dataList.add(modelType);
                } catch (InstantiationException e) {
                } catch (IllegalAccessException e) {
                }
            }
            curCell = 0;
            line++;
        } else if (name.equals("c")) { // c => cell

            if(titleType.equals(TitleType.SIMPLE)){
                if(line == start){
                    String cellStr = attributes.getValue("r");
                    titleString = cellStr.replaceAll("[0-9]", "");
                }
            }

            if (line > start) {
                String cellStr = attributes.getValue("r");
                String key = cellStr.replaceAll("[0-9]", "");

                if(titleType.equals(TitleType.SIMPLE)){
                    key = titlemap.get(key);
                }

                Map<String, String> confMap = ExcelUtil.getExcelCells(clazz);
                for (Map.Entry<String, String> entry : confMap.entrySet()) {
                    if (StringUtils.endsWithIgnoreCase(key, entry.getValue())) {
                        fieldmap.put(curCell, entry.getKey());
                    }
                }
                String cellType = attributes.getValue("t");
                if (cellType != null && cellType.equals("s")) {
                    nextIsString = true;
                } else {
                    nextIsString = false;
                }
            }
        }
        lastContents = "";
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {

        if(titleType.equals(TitleType.SIMPLE)){
            if(line==start){
                int idx = Integer.parseInt(lastContents);
                String title = new XSSFRichTextString(sst.getEntryAt(idx)).toString().trim();
                titlemap.put(titleString, title);
            }
        }

        if (line > start) {
            if (nextIsString) {
                int idx = Integer.parseInt(lastContents);
                lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString().trim();
                nextIsString = false;
            }
            if (name.equals("v")) {
                String fieldName = fieldmap.get(curCell);
                if (StringUtils.isNotBlank(fieldName)) {
                    String setMethodName = "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                    try {
                        Field field = FieldUtils.getField(clazz, fieldName, true);
                        Method setMethod = clazz.getMethod(setMethodName, new Class[] { field.getType() });
                        Type[] ts = setMethod.getGenericParameterTypes();
                        // 只要一个参数,判断类型
                        String xclass = ts[0].toString();
                        setValue(xclass, modelType, setMethod, lastContents);
                    } catch (Exception e) {
                    }
                }
                curCell++;
            }
        }
    }
    public void setValue(String xclass, ModelType tObject, Method setMethod, String value) throws Exception {
        if (xclass.equals("class java.lang.String")) {
            setMethod.invoke(tObject, value);
        } else if (xclass.equals("class java.util.Date")) {
            setMethod.invoke(tObject, DateUtil.getJavaDate(new Double(value), false));
        } else if (xclass.equals("class java.lang.Boolean")) {
            Boolean boolname = true;
            if (value.equals("否")) {
                boolname = false;
            }
            setMethod.invoke(tObject, boolname);
        } else if (xclass.equals("class java.lang.Short")) {
            setMethod.invoke(tObject, new Short(value));
        } else if (xclass.equals("class java.lang.Integer")) {
            setMethod.invoke(tObject, new Integer(value));
        } else if (xclass.equals("class java.lang.Long")) {
            setMethod.invoke(tObject, new Long(value));
        } else if (xclass.equals("class java.lang.Double")) {
            setMethod.invoke(tObject, new Double(value));
        } else if (xclass.equals("class java.math.BigDecimal")) {
            setMethod.invoke(tObject, new BigDecimal(value));
        } else {
            setMethod.invoke(tObject, value);
        }
    }

    public boolean isTitleLine() {
        return line < start;
    }

    @Override
    public void characters(char[] ch, int start, int length) throws SAXException {
        lastContents += new String(ch, start, length);
    }

    public List<ModelType> getData() {
        return dataList;
    }



}