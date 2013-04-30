package com.catapult.excel.parsing;

import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelWorkbook {

    public Map<String, ExcelSheet> excelSheetMap = new LinkedHashMap<String, ExcelSheet>(0);

    public Map<String, ExcelSheet> getExcelSheetMap() {
        return excelSheetMap;
    }

    public void setExcelSheetMap(Map<String, ExcelSheet> excelSheetMap) {
        this.excelSheetMap = excelSheetMap;
    }

}
