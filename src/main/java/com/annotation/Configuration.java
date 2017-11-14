package com.annotation;

import java.util.Map;

/**
 * Created by user on 2017/11/14.
 */
public interface Configuration {
    Map<String, String> parse(Class<?> clazz) throws Exception;
}
