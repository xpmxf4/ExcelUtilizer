package com.example.excelutilizer.v2.dto;

import com.example.excelutilizer.v2.annotation.ExcelChapter;
import com.example.excelutilizer.v2.annotation.ExcelHyperLink;
import com.fasterxml.jackson.annotation.JsonInclude;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 필드 또는 메서드의 메타데이터를 저장하는 DTO
 */
@Getter
@AllArgsConstructor
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ExcelSupport {

    private String name;  // 필드 또는 메서드 이름
    private int order = Integer.MAX_VALUE;  // 순서
    private String chapterTitle;  // 챕터 제목
    private Field field;  // 반영된 필드
    private Method method;  // 반영된 메서드
    private boolean usedInSheet;  // 시트에 사용되는지 여부
    private boolean isUsedHyperlink;  // 하이퍼링크 사용 여부
    private String color;  // 색상 코드
    private boolean isHangulField;  // 한글 필드 여부
    private SupportType supportType;  // 지원 타입 (필드 또는 메서드)

    // 필드에서 ExcelSupport 객체를 생성하는 메서드
    public static ExcelSupport fromField(Field field) {
        field.setAccessible(true);
        ExcelChapter chapterAnnotation = field.getAnnotation(ExcelChapter.class);
        return new ExcelSupport(
            field.getName(),
            chapterAnnotation != null ? chapterAnnotation.order() : Integer.MAX_VALUE,
            chapterAnnotation != null ? chapterAnnotation.value() : field.getName(),
            field,
            null,
            chapterAnnotation != null,
            chapterAnnotation != null && field.isAnnotationPresent(ExcelHyperLink.class),
            chapterAnnotation != null ? chapterAnnotation.color() : "",
            chapterAnnotation != null && chapterAnnotation.isHangulField(),
            SupportType.FIELD
        );
    }

    // 메서드에서 ExcelSupport 객체를 생성하는 메서드
    public static ExcelSupport fromMethod(Method method) {
        method.setAccessible(true);
        ExcelChapter chapterannotation = method.getAnnotation(ExcelChapter.class);
        return new ExcelSupport(
            method.getName(),
            chapterannotation != null ? chapterannotation.order() : Integer.MAX_VALUE,
            chapterannotation != null ? chapterannotation.value() : method.getName(),
            null,
            method,
            chapterannotation != null,
            chapterannotation != null && method.isAnnotationPresent(ExcelHyperLink.class),
            chapterannotation != null ? chapterannotation.color() : "",
            chapterannotation != null && chapterannotation.isHangulField(),
            SupportType.METHOD
        );
    }

    public enum SupportType {
        FIELD, METHOD
    }

    public boolean isField() {
        return supportType == SupportType.FIELD;
    }

    public boolean isMethod() {
        return supportType == SupportType.METHOD;
    }
}
