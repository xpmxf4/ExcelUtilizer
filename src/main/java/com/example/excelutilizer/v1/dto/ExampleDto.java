package com.example.excelutilizer.v1.dto;

import com.example.excelutilizer.v1.annotation.ExcelColumn;
import lombok.Data;

@Data
public class ExampleDto {

    @ExcelColumn(name = "name", order = 1, color = "#341ABC", isHangulField = true)
    private String name;

    @ExcelColumn(name = "email", order = 2, color = "#341ABC", isHangulField = false)
    private String email;

}
