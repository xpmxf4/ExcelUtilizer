package com.example.excelutilizer.v2.dto;

import com.example.excelutilizer.v2.annotation.ExcelChapter;
import com.example.excelutilizer.v2.annotation.ExcelHyperLink;
import lombok.Getter;

@Getter
public class ExampleDto {

    @ExcelChapter(value = "이름", order = 1, color = "#FF0000", isHangulField = true)
    private String name;

    @ExcelChapter(value = "Age", order = 2, isHangulField = false)
    private int age;

    @ExcelChapter(value = "Profile Link", order = 3, isHangulField = false)
    @ExcelHyperLink("http://naver.com")
    private String profileLink;
}
