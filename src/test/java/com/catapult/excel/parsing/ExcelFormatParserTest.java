package com.catapult.excel.parsing;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.IteratorUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

public class ExcelFormatParserTest {

    @SuppressWarnings("unchecked")
    @Test
    public void extractExcelFormat() throws InvalidFormatException, IOException{
        //String filePath = "D:/Development/workspace/final/qms/new-qms-web-ph/src/test/java/test/department/format/parsing/department-setup-template.xlsx";
//        String filePath = "D:/Development/workspace/final/qms/new-qms-web-ph/src/test/java/test/department/format/parsing/format-test.xlsx";
//        File file = new File(filePath);

        InputStream is = getClass().getResourceAsStream("format-test.xlsx");
        Workbook workbook = WorkbookFactory.create(is);
        is.close();

        Map<String, String> cellValueMap = new HashMap<String, String>(0);
        Map<String, CellStyle> cellStyleMap = new HashMap<String, CellStyle>(0);
        String key = null;
        CellStyle cellStyle = null;
        String cellValue = null;
        int formatPoints = 0;
        ExcelHeader excelHeader = null;

        Map<String, ExcelHeader> regularCellList = new LinkedHashMap<String, ExcelHeader>(0);
        Map<String, ExcelHeader> mergedCellList = new LinkedHashMap<String, ExcelHeader>(0);
        Map<String, ExcelHeader> finalCellList = new LinkedHashMap<String, ExcelHeader>(0);

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            List<Row> rowList = (List<Row>) IteratorUtils.toList(sheet.iterator());
            List<Cell> cellList = null;

            // regular cells
            for(Row row : rowList){
                cellList = (List<Cell>) IteratorUtils.toList(row.iterator());

                for(Cell cell : cellList){
                    key = String.valueOf(sheetIndex) +','+ String.valueOf(cell.getRowIndex()) +','+ String.valueOf(cell.getColumnIndex());

                    cellStyle = cell.getCellStyle();
                    if(cellStyle.getBorderBottom() != CellStyle.BORDER_NONE){
                        cellStyleMap.put(key, cellStyle);
                    }

                    cellValue = cell.toString();
                    if(cellStyleMap.containsKey(key) && StringUtils.isNotBlank(cellValue)){
                        cellValueMap.put(key, cellValue);
                    }


                    formatPoints = 0;


                    if(cellStyle.getBorderLeft() != CellStyle.BORDER_NONE){
                        formatPoints++;
                    }

                    if(cellStyle.getBorderRight() != CellStyle.BORDER_NONE){
                        formatPoints++;
                    }

                    if(cellStyle.getBorderTop() != CellStyle.BORDER_NONE){
                        formatPoints++;
                    }

                    if(cellStyle.getBorderBottom() != CellStyle.BORDER_NONE){
                        formatPoints++;
                    }

                    if(cellStyle.getAlignment() != CellStyle.ALIGN_GENERAL){
                        formatPoints++;
                    }

                    if(cellStyle.getFillBackgroundColor() != CellStyle.NO_FILL){
                        formatPoints++;
                    }

                    if(cellStyle.getFillForegroundColor() != CellStyle.NO_FILL){
                        formatPoints++;
                    }

                    if(formatPoints > 1){
                        excelHeader = new ExcelHeader();
                        excelHeader.setCell(cell);
                        excelHeader.setCellRangeAddress(null);
                        excelHeader.setFormatPoints(formatPoints);
                        excelHeader.setMergedCell(false);
                        regularCellList.put(key, excelHeader);
                        System.out.println(formatPoints);
                    }
                }
            }

            // merge cells
            Cell cell = null;
            int rowIndex = 0;
            int columnIndex = 0;
            for(int i = 0; i < sheet.getNumMergedRegions(); i++){
                CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);

                rowIndex = cellRangeAddress.getFirstRow();
                columnIndex = cellRangeAddress.getFirstColumn();
                cell = sheet.getRow(rowIndex).getCell(columnIndex);

                key = String.valueOf(sheetIndex) +','+ String.valueOf(cell.getRowIndex()) +','+ String.valueOf(cell.getColumnIndex());

                cellStyle = cell.getCellStyle();
                if(cellStyle.getBorderBottom() != CellStyle.BORDER_NONE){
                    cellStyleMap.put(key, cellStyle);
                }

                cellValue = cell.toString();
                if(cellStyleMap.containsKey(key) && StringUtils.isNotBlank(cellValue)){
                    cellValueMap.put(key, cellValue);
                }


                formatPoints = 0;


                if(cellStyle.getBorderLeft() != CellStyle.BORDER_NONE){
                    formatPoints++;
                }

                if(cellStyle.getBorderRight() != CellStyle.BORDER_NONE){
                    formatPoints++;
                }

                if(cellStyle.getBorderTop() != CellStyle.BORDER_NONE){
                    formatPoints++;
                }

                if(cellStyle.getBorderBottom() != CellStyle.BORDER_NONE){
                    formatPoints++;
                }

                if(cellStyle.getAlignment() != CellStyle.ALIGN_GENERAL){
                    formatPoints++;
                }

                if(cellStyle.getFillBackgroundColor() != CellStyle.NO_FILL){
                    formatPoints++;
                }

                if(cellStyle.getFillForegroundColor() != CellStyle.NO_FILL){
                    formatPoints++;
                }

                if(formatPoints > 1){
                    excelHeader = new ExcelHeader();
                    excelHeader.setCell(cell);
                    excelHeader.setCellRangeAddress(cellRangeAddress);
                    excelHeader.setFormatPoints(formatPoints);
                    excelHeader.setMergedCell(true);
                    if(excelHeader != null){
                        mergedCellList.put(key, excelHeader);
                    }
                    System.out.println(formatPoints);
                }
            }

            // remove merge cells from regular cell list

        }

        System.out.println(cellValueMap.toString());
//        System.out.println(cellValueMap.size());
//        System.out.println(cellStyleMap.toString());
//        System.out.println(cellStyleMap.size());
//        String filePathCopy = "C:/Users/jcosare/Desktop/tmp/department-setup-template.xlsx";
//        File fileCopy = new File(filePathCopy);
//        Workbook workbookCopy = WorkbookFactory.create(fileCopy);
//
//        key = null;
//        cellStyle = null;
//        cellValue = null;
//
//        for(int sheetIndex = 0; sheetIndex != -1; sheetIndex++){
//            Sheet sheet = workbookCopy.createSheet();
//            for(int rowIndex = 0;;rowIndex++){
//                Row row = sheet.createRow(rowIndex);
//
//                for(int columnIndex = 0;;columnIndex++){
//                    key = String.valueOf(sheetIndex) +','+ String.valueOf(rowIndex) +','+ String.valueOf(columnIndex);
//                    if(cellStyleMap.get(key) != null){
//                        cellStyle = cellStyleMap.get(key);
//                        cellValue = cellValueMap.get(key);
//
//                        Cell cell = row.createCell(columnIndex);
//                        cell.setCellStyle(cellStyle);
//                        cell.setCellValue(cellValue);
//                    }else{
//                        break;
//                    }
//                }
//            }
//        }

    }
}
