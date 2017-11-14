package com.annotation;

import com.google.common.collect.Maps;
import org.springframework.beans.BeanUtils;
import org.springframework.util.ReflectionUtils;

import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * Created by user on 2017/11/14.
 */
public class AnnotationConfiguration implements Configuration {
    @Override
    public Map<String, String> parse(Class<?> clazz) throws Exception {
        Map<String, String> confMap = Maps.newHashMap();
        PropertyDescriptor[] propertyDescriptors = BeanUtils.getPropertyDescriptors(clazz);
        for (PropertyDescriptor propertyDescriptor : propertyDescriptors) {
            Field field = ReflectionUtils.findField(clazz, propertyDescriptor.getName());
            if (null != field && field.isAnnotationPresent(ExcelCell.class)) {
                ExcelCell cell = field.getAnnotation(ExcelCell.class);
                confMap.put(field.getName(),cell.value());
            }
        }
        return confMap;
    }
}
