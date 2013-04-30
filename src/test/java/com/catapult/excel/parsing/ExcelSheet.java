package com.catapult.excel.parsing;

import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelSheet {

    public Map<String, ExcelSection> excelSectionMap = new LinkedHashMap<String, ExcelSection>(0);

    public Map<String, ExcelSection> getExcelSectionMap() {
        return excelSectionMap;
    }

    public void setExcelSectionMap(Map<String, ExcelSection> excelSectionMap) {
        this.excelSectionMap = excelSectionMap;
    }

}
