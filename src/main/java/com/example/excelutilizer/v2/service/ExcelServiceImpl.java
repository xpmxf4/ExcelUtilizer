package com.example.excelutilizer.v2.service;

import com.example.excelutilizer.v2.annotation.ExcelHyperLink;
import com.example.excelutilizer.v2.dto.ExcelSupport;
import java.awt.Color;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelServiceImpl implements ExcelService {

    private final String defaultFontName = "맑은 고딕";

    @Override
    public <T> Workbook create(String sheetTitle, List<T> contents, Class<T> clazz) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFCreationHelper creationHelper = workbook.getCreationHelper();

        List<ExcelSupport> excelSupports = Stream.concat(
                Stream.of(clazz.getDeclaredFields()).map(ExcelSupport::fromField),
                Stream.of(clazz.getDeclaredMethods()).map(ExcelSupport::fromMethod)
            )
            .sorted(Comparator.comparingInt(ExcelSupport::getOrder))
            .toList();

        List<ExcelSupport> filteredSupports = excelSupports.stream()
            .filter(ExcelSupport::isUsedInSheet)
            .collect(Collectors.toList());

        Map<String, ExcelSupport> excelSupportMap = excelSupports.stream()
            .collect(Collectors.toMap(ExcelSupport::getName, support -> support));

        Sheet sheet = setupSheet(workbook, sheetTitle);
        setupHeaderRow(sheet, filteredSupports, workbook, creationHelper);
        fillSheetContents(sheet, contents, filteredSupports, excelSupportMap, creationHelper);

        autoSizeColumns(sheet, filteredSupports);
        return workbook;
    }

    private void autoSizeColumns(Sheet sheet, List<ExcelSupport> filteredSupports) {
        for (int index = 0; index < filteredSupports.size(); index++) {
            int originColumnWidth = sheet.getColumnWidth(index);
            if (filteredSupports.get(index).isHangulField()) {
                sheet.autoSizeColumn(index);
                int autoSizeColumnWidth = 9 * sheet.getColumnWidth(index) / 5;
                sheet.setColumnWidth(index, Math.max(autoSizeColumnWidth, originColumnWidth));
            } else {
                sheet.autoSizeColumn(index);
                sheet.setColumnWidth(index, Math.max(sheet.getColumnWidth(index), originColumnWidth));
            }
        }
    }

    private <T> void fillSheetContents(Sheet sheet, List<T> contents, List<ExcelSupport> filteredSupports, Map<String, ExcelSupport> excelSupportMap, XSSFCreationHelper creationHelper)
        throws InvocationTargetException, IllegalAccessException {
        CellStyle baseCellStyle = createBaseCellStyle(sheet.getWorkbook());
        CellStyle hyperLinkStyle = createHyperLinkStyle(sheet.getWorkbook(), baseCellStyle);

        for (int rowIndex = 0; rowIndex < contents.size(); rowIndex++) {
            T content = contents.get(rowIndex);
            Row row = sheet.createRow(rowIndex + 1);

            for (int columnIndex = 0; columnIndex < filteredSupports.size(); columnIndex++) {
                ExcelSupport support = filteredSupports.get(columnIndex);
                Cell cell = row.createCell(columnIndex);
                cell.setCellStyle(baseCellStyle);

                String cellValue = support.isMethod() ? invokeMethod(content, support.getMethod()) :
                    support.isField() ? getObjectValueAsString(support.getField().get(content)) : "";

                cell.setCellValue(cellValue);

                if (support.isUsedHyperlink() && !cellValue.isEmpty()) {
                    String linkValue = processHyperlinkValue(content, support, cellValue, excelSupportMap);
                    if (!linkValue.isBlank()) {
                        Hyperlink link = creationHelper.createHyperlink(HyperlinkType.URL);
                        link.setAddress(linkValue);
                        cell.setHyperlink(link);
                    }
                    cell.setCellStyle(hyperLinkStyle);
                }
            }
        }
    }

    private <T> String processHyperlinkValue(T content, ExcelSupport support, String cellValue, Map<String, ExcelSupport> excelSupportMap) throws InvocationTargetException, IllegalAccessException {
        ExcelHyperLink linkAnnocation = support.isMethod() ? support.getMethod().getAnnotation(ExcelHyperLink.class) :
            support.isField() ? support.getField().getAnnotation(ExcelHyperLink.class) : null;

        if (linkAnnocation == null) {
            return cellValue;
        }

        String value = linkAnnocation.value();

        if (value.isEmpty()) {
            return cellValue;
        }

        Pattern fieldPattern = Pattern.compile("<<([^>]+)>>");
        Matcher fieldMatcher = fieldPattern.matcher(value);
        while (fieldMatcher.find()) {
            String fieldName = fieldMatcher.group(1);
            String fieldValue = excelSupportMap.containsKey(fieldName) ? getObjectValueAsString(excelSupportMap.get(fieldName).getField().get(content)) : "";
            value = value.replace("<<" + fieldName + ">>", fieldValue);
        }

        Pattern methodPattern = Pattern.compile("\\{\\{([^}]+)\\}\\}");
        Matcher methodMatcher = methodPattern.matcher(value);
        while (methodMatcher.find()) {
            String methodName = methodMatcher.group(1);
            Method method = excelSupportMap.get(methodName).getMethod();
            String methodResult = method != null ? method.invoke(content).toString() : "";
            value = value.replace("{{" + methodName + "}}", methodResult);
        }

        return value;
    }

    private String getObjectValueAsString(Object o) {
        return null;
    }

    private <T> String invokeMethod(T content, Method method) {
        try {
            method.setAccessible(true);
            return method.invoke(content).toString();
        } catch (Exception e) {
            System.err.println(e.getMessage());
            return "";
        }
    }

    private CellStyle createHyperLinkStyle(Workbook workbook, CellStyle baseCellStyle) {
        Font hyperlinkFont = workbook.createFont();
        hyperlinkFont.setFontName(defaultFontName);
        hyperlinkFont.setColor(IndexedColors.BLUE.getIndex());
        hyperlinkFont.setUnderline(Font.U_SINGLE);

        CellStyle hyperlinkStyle = workbook.createCellStyle();
        hyperlinkStyle.cloneStyleFrom(baseCellStyle);
        hyperlinkStyle.setFont(hyperlinkFont);
        return hyperlinkStyle;
    }

    private CellStyle createBaseCellStyle(Workbook workbook) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName(defaultFontName);

        CellStyle baseCellStyle = workbook.createCellStyle();
        baseCellStyle.setFont(baseFont);
        baseCellStyle.setAlignment(HorizontalAlignment.LEFT);
        baseCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        baseCellStyle.setBorderTop(BorderStyle.THIN);
        baseCellStyle.setBorderBottom(BorderStyle.THIN);
        baseCellStyle.setBorderLeft(BorderStyle.THIN);
        baseCellStyle.setBorderRight(BorderStyle.THIN);
        return baseCellStyle;
    }


    // sheet 설정
    private Sheet setupSheet(XSSFWorkbook workbook, String sheetTitle) {
        Sheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(sheetTitle));
        sheet.setHorizontallyCenter(true);
        sheet.setHorizontallyCenter(true);
        return sheet;
    }

    private void setupHeaderRow(Sheet sheet, List<ExcelSupport> filteredSupports, XSSFWorkbook workbook, XSSFCreationHelper creationHelper) {
        CellStyle headerStyle = createHeaderStyle(workbook);
        Row headerRow = sheet.createRow(0);

        // 헤더 스타일 설정
        for (int i = 0; i < filteredSupports.size(); i++) {
            ExcelSupport support = filteredSupports.get(i);
            Cell cell = headerRow.getCell(i);
            cell.setCellValue(support.getChapterTitle());
            cell.setCellStyle(determineCellStyle(workbook, support, headerStyle, creationHelper));
        }

        // 헤더 사이즈 설정
        autoHeaderSizeColumns(sheet, filteredSupports);
    }

    private void autoHeaderSizeColumns(Sheet sheet, List<ExcelSupport> filteredSupports) {
        for (int i = 0; i < filteredSupports.size(); i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, 9 * sheet.getColumnWidth(i) / 5);           // 왜 column width 가 이렇게 되어있을까
        }
    }

    // 셀 스타일 결정
    private CellStyle determineCellStyle(XSSFWorkbook workbook, ExcelSupport support, CellStyle baseStyle, XSSFCreationHelper creationHelper) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.cloneStyleFrom(baseStyle);
        if (support.getColor() != null && !support.getColor().isEmpty()) {
            XSSFColor color = new XSSFColor(Color.decode(support.getColor()), new DefaultIndexedColorMap());
            cellStyle.setFillForegroundColor(color.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        return cellStyle;
    }

    // 기본 헤더 스타일 생성
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setFontName(defaultFontName);
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);

        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        return headerStyle;
    }


}
