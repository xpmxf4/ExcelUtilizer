package com.example.excelutilizer.v1.enums;

import java.util.EnumMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.util.Map;

public enum CellTypeHandler {
    STRING(CellType.STRING) {
        @Override
        public String getCellValue(Cell cell) {
            return cell.getStringCellValue();
        }
    },
    NUMERIC(CellType.NUMERIC) {
        @Override
        public String getCellValue(Cell cell) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toString();
            } else {
                return String.valueOf(cell.getNumericCellValue());
            }
        }
    },
    BOOLEAN(CellType.BOOLEAN) {
        @Override
        public String getCellValue(Cell cell) {
            return String.valueOf(cell.getBooleanCellValue());
        }
    },
    FORMULA(CellType.FORMULA) {
        @Override
        public String getCellValue(Cell cell) {
            return cell.getCellFormula();
        }
    },
    BLANK(CellType.BLANK) {
        @Override
        public String getCellValue(Cell cell) {
            return "";
        }
    },
    ERROR(CellType.ERROR) {
        @Override
        public String getCellValue(Cell cell) {
            return String.valueOf(cell.getErrorCellValue());
        }
    };

    private static final Map<CellType, CellTypeHandler> MAP = new EnumMap<>(CellType.class);

    static {
        for (CellTypeHandler handler : values()) {
            MAP.put(handler.cellType, handler);
        }
    }

    private final CellType cellType;

    CellTypeHandler(CellType cellType) {
        this.cellType = cellType;
    }

    public abstract String getCellValue(Cell cell);

    public static CellTypeHandler fromCell(Cell cell) {
        CellTypeHandler handler = MAP.get(cell.getCellType());
        if (handler == null) {
            throw new IllegalArgumentException("Unknown cell type: " + cell.getCellType());
        }
        return handler;
    }
}
