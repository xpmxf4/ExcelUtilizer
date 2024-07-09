package com.example.excelutilizer.v1.annotation;

import static java.lang.annotation.RetentionPolicy.*;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

@Retention(RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelColumn {
    String name();
    String color();
    int order();
    boolean isHangulField();
}
