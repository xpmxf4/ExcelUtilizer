package com.example.excelutilizer.v2.annotation;

import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

/**
 * Excel 시트의 챕터를 정의하는 어노테이션
 */
@Retention(RUNTIME)
@Target({ElementType.FIELD, ElementType.METHOD})
public @interface ExcelChapter {

    String value();

    String color() default "";

    int order() default Integer.MAX_VALUE;

    boolean isHangulField();
}
