package com.catapult.excel.parsing;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelSection {

    public static final int DEFAULT_VALUE = -1;

    public static final int ORIENTATION_HORIZONTAL = 1;

    public static final int ORIENTATION_VERTICAL = 2;

    public static final int CELL_SPACE_WITHIN = 1;

    public static final int CELL_SPACE_VERTICAL = 2;

    public static final int CELL_SPACE_HORIZONTAL = 3;

    public static final int CELL_SPACE_CROSS = 4;

    private boolean isSet = false;

    private boolean isOrientationFixed = false;

    private boolean isOrientationRewindNeeded = false;

    private final String hashKey = RandomStringUtils.randomAlphanumeric(16);

    private Map<String, String> dataMap = new HashMap<String, String>(0);

    private Map<String, String> headerMap = new HashMap<String, String>(0);

    private CellRangeAddress sectionCellRange = new CellRangeAddress(DEFAULT_VALUE,DEFAULT_VALUE,DEFAULT_VALUE,DEFAULT_VALUE);

    private CellRangeAddress headerCellRange = new CellRangeAddress(DEFAULT_VALUE,DEFAULT_VALUE,DEFAULT_VALUE,DEFAULT_VALUE);

    private CellRangeAddress dataCellRange = new CellRangeAddress(DEFAULT_VALUE,DEFAULT_VALUE,DEFAULT_VALUE,DEFAULT_VALUE);

    private int headerOrientation = DEFAULT_VALUE;

    private int dataOrientation = DEFAULT_VALUE;

    public int getHeaderOrientation() {
        return headerOrientation;
    }

    public void setHeaderOrientation(int headerOrientation) {
        this.headerOrientation = headerOrientation;
    }

    public int getDataOrientation() {
        return dataOrientation;
    }

    public void setDataOrientation(int dataOrientation) {
        this.dataOrientation = dataOrientation;
    }

    public Map<String, String> getDataMap() {
        return dataMap;
    }

    public void setDataMap(Map<String, String> dataMap) {
        this.dataMap = dataMap;
    }

    public Map<String, String> getHeaderMap() {
        return headerMap;
    }

    public void setHeaderMap(Map<String, String> headerMap) {
        this.headerMap = headerMap;
    }

    public CellRangeAddress getSectionCellRange() {
        return sectionCellRange;
    }

    public void setSectionCellRange(CellRangeAddress sectionCellRange) {
        this.sectionCellRange = sectionCellRange;
    }

    public CellRangeAddress getHeaderCellRange() {
        return headerCellRange;
    }

    public void setHeaderCellRange(CellRangeAddress headerCellRange) {
        this.headerCellRange = headerCellRange;
    }

    public CellRangeAddress getDataCellRange() {
        return dataCellRange;
    }

    public void setDataCellRange(CellRangeAddress dataCellRange) {
        this.dataCellRange = dataCellRange;
    }

    public String getHashKey() {
        return hashKey;
    }

    public boolean isSet() {
        return isSet;
    }

    public void setSet(boolean isSet) {
        this.isSet = isSet;
    }

    public boolean isOrientationFixed() {
        return isOrientationFixed;
    }

    public void setOrientationFixed(boolean isOrientationFixed) {
        this.isOrientationFixed = isOrientationFixed;
    }

    public boolean isOrientationRewindNeeded() {
        return isOrientationRewindNeeded;
    }

    public void setOrientationRewindNeeded(boolean isOrientationRewindNeeded) {
        this.isOrientationRewindNeeded = isOrientationRewindNeeded;
    }

}
