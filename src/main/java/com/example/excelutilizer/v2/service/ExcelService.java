package com.example.excelutilizer.v2.service;

import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;

public interface ExcelService {

    <T>Workbook create(String sheetTitle, List<T> content, Class<T> tc) throws Exception;
}
