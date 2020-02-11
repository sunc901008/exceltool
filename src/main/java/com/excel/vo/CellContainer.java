package com.excel.vo;

import com.excel.base.Common;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.text.DecimalFormat;
import java.text.ParseException;

/**
 * @author sunc
 * @date 2019/11/14 8:54
 * @description CellContainer
 */

public class CellContainer {

    public static final String STRING = "string";
    public static final String FORMULA = "formula";
    public static final String NUMBER = "number";
    public static final String BOOLEAN = "boolean";
    public static final String DATE = "date";
    public static final String NONE = "none";

    private String type;
    private String value;

    public CellContainer(String type, String value) {
        this.type = type;
        this.value = value;
    }

    public void setCell(Cell cell) {
        switch (this.type) {
            case STRING:
                cell.setCellType(CellType.STRING);
                cell.setCellValue(value);
                break;
            case FORMULA:
                cell.setCellType(CellType.FORMULA);
                cell.setCellFormula(value);
                break;
            case NUMBER:
                cell.setCellType(CellType.NUMERIC);
                Double cellValue = transfer(value);
                if (cellValue == null) {
                    cell.setCellValue(value);
                } else {
                    cell.setCellValue(cellValue);
                }
                break;
            case BOOLEAN:
                cell.setCellType(CellType.BOOLEAN);
                cell.setCellValue(Boolean.parseBoolean(value));
                break;
            case DATE:
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(Common.datetimeParse(value));
                break;
            default:
                cell.setCellType(CellType.BLANK);
                cell.setCellValue("");
        }
    }

    private Double transfer(String value) {
        try {
            if (value.contains("%")) {
                if (value.contains(".")) {
                    Number v = new DecimalFormat("##.00%").parse(value);
                    return v.doubleValue();
                } else {
                    Number v = new DecimalFormat("##%").parse(value);
                    return v.doubleValue();

                }
            } else if (value.contains(".")) {
                return Double.parseDouble(value);
            } else {
                return (double) Integer.parseInt(value);
            }
        } catch (ParseException e) {
            e.printStackTrace();
            return null;
        }
    }

}
