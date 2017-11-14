package com.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by user on 2017/11/13.
 */

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Excel {
    /**
     * 从多少行开始解析数据
     *
     * @return
     */
    int start() default 1;

    Struct struct() default Struct.ANNOTATION;

    TitleType titleType() default TitleType.SIMPLE;

    enum Struct {
        ANNOTATION, XML
    }

    enum TitleType {
        SIMPLE, MULTIPLE
    }
}

