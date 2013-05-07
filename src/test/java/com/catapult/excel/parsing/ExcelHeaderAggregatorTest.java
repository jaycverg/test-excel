package com.catapult.excel.parsing;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

public class ExcelHeaderAggregatorTest
{

    @Test
    public void extractExcelFormat() throws InvalidFormatException, IOException
    {
//        String filePath = "D:/Development/workspace/final/qms/new-qms-web-ph/src/test/java/test/department/format/parsing/department-setup-template.xlsx";
//        String filePath = "D:/Development/workspace/final/qms/new-qms-web-ph/src/test/java/test/department/format/parsing/format-test.xlsx";
//        String filePath = "C:/Users/jcosare/Downloads/HWL CUSTOMER 2 TEMPLATE - AIR FREIGHT.xls";
//        File file = new File(filePath);

        InputStream is = getClass().getResourceAsStream("format-test.xlsx");
        Workbook workbook = WorkbookFactory.create(is);
        is.close();

        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        // regular cells
        String key = null;
        String cellValue = null;

        // flattened cells (regular + merge cells)
        ExcelWorkbook currentExcelWorkbook = new ExcelWorkbook();
        ExcelSheet currentExcelSheet = null;
        ExcelSection currentExcelSection = null;

        Map<String, String> cellValueMap = new HashMap<String, String>(0);

        // iterate all sheets
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            sheet = workbook.getSheetAt(sheetIndex);
            currentExcelSheet = new ExcelSheet();
            currentExcelSection = this.getExcelSectionBasedOnSheet(sheet, sheetIndex, workbook, currentExcelSheet.getExcelSectionMap(), currentExcelSection);
            if (!workbook.isSheetHidden(sheetIndex)) {
                // iterating sheet rows
                // regular cells
                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    row = sheet.getRow(rowIndex);
                    currentExcelSection = this.getExcelSectionBasedOnRow(row, sheetIndex, currentExcelSheet.getExcelSectionMap(), currentExcelSection);
                    if (row != null) {
                        // iterating sheet row columns
                        for (int columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {
                            cell = row.getCell(columnIndex);
                            currentExcelSection = this.getExcelSectionBasedOnCell(cell, sheet, rowIndex, columnIndex, currentExcelSheet.getExcelSectionMap(), currentExcelSection);
                            if (cell != null) {
                                this.setCellValue(cell, sheetIndex, cellValueMap);
                                currentExcelSection = this.setBoxCoordinates(cell, sheetIndex, currentExcelSection);
                                this.setExcelSectionMap(currentExcelSheet.getExcelSectionMap(), currentExcelSection);
                            }

                            if (currentExcelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_VERTICAL
                                    && currentExcelSection.isOrientationFixed()
                                    && currentExcelSection.isOrientationRewindNeeded()) {

                                int rewindedColumnIndex = currentExcelSection.getHeaderCellRange().getFirstColumn();
                                currentExcelSection.setHeaderCellRange(this.resetCellRangeAddress(currentExcelSection.getHeaderCellRange()));
                                currentExcelSection.setDataCellRange(this.resetCellRangeAddress(currentExcelSection.getDataCellRange()));
                                currentExcelSection.setSectionCellRange(this.resetCellRangeAddress(currentExcelSection.getSectionCellRange()));

                                while (rewindedColumnIndex <= columnIndex) {
                                    cell = row.getCell(rewindedColumnIndex);
                                    currentExcelSection = this.getExcelSectionBasedOnCell(cell, sheet, rowIndex, columnIndex, currentExcelSheet.getExcelSectionMap(), currentExcelSection);
                                    if (cell != null) {
                                        this.setCellValue(cell, sheetIndex, cellValueMap);
                                        currentExcelSection = this.setBoxCoordinates(cell, sheetIndex, currentExcelSection);
                                        this.setExcelSectionMap(currentExcelSheet.getExcelSectionMap(), currentExcelSection);
                                    }
                                    rewindedColumnIndex++;
                                }
                                currentExcelSection.setOrientationRewindNeeded(false);
                            }
                        }
                    }
                }

                // merged cells
                for (int mergeIndex = 0; mergeIndex < sheet.getNumMergedRegions(); mergeIndex++) {
                    CellRangeAddress cellRangeAddress = sheet.getMergedRegion(mergeIndex);
                    this.populateMergedCellData(sheet, sheetIndex, cellValueMap, cellRangeAddress);
                }

                // remove empty excel sheet
//                Collection<ExcelSection> excelSectionList = (Collection<ExcelSection>) currentExcelSheet.getExcelSectionMap().values();
//                for(ExcelSection excelSectionItem : excelSectionList){
//                    if(this.isCellRangeNull(excelSectionItem.getHeaderCellRange())
//                        && this.isCellRangeNull(excelSectionItem.getDataCellRange())
//                        && this.isCellRangeNull(excelSectionItem.getSectionCellRange())){
//                        currentExcelSheet.getExcelSectionMap().remove(excelSectionItem.getHashKey());
//                    }
//                }

                // flattened cells (regular + merge cells)
                System.out.println("=====================================");
                System.out.println(sheet.getSheetName());
                System.out.println("=====================================");
                for (ExcelSection excelSectionItem : currentExcelSheet.getExcelSectionMap().values()) {

                    if (this.isCellRangeNotNull(excelSectionItem.getHeaderCellRange())
                            && this.isCellRangeNotNull(excelSectionItem.getDataCellRange())
                            && this.isCellRangeNotNull(excelSectionItem.getSectionCellRange())) {
                        excelSectionItem = this.processCellValueMap(sheetIndex, cellValueMap, excelSectionItem);
                        System.out.println("Header Cell Range : " + excelSectionItem.getHeaderCellRange().getFirstRow() + " , " + excelSectionItem.getHeaderCellRange().getFirstColumn() + " , " + excelSectionItem.getHeaderCellRange().getLastRow() + " , " + excelSectionItem.getHeaderCellRange().getLastColumn());
                        System.out.println("Data Cell Range : " + excelSectionItem.getDataCellRange().getFirstRow() + " , " + excelSectionItem.getDataCellRange().getFirstColumn() + " , " + excelSectionItem.getDataCellRange().getLastRow() + " , " + excelSectionItem.getDataCellRange().getLastColumn());
                        System.out.println("Section Cell Range : " + excelSectionItem.getSectionCellRange().getFirstRow() + " , " + excelSectionItem.getSectionCellRange().getFirstColumn() + " , " + excelSectionItem.getSectionCellRange().getLastRow() + " , " + excelSectionItem.getSectionCellRange().getLastColumn());
                        for (String value : excelSectionItem.getHeaderMap().values()) {
                            System.out.println(value);
                        }

                        int dataNewLine = excelSectionItem.getHeaderMap().size();

                        int i = 0;
                        System.out.println("*/////////////////////////////*");
                        for (String value : excelSectionItem.getDataMap().values()) {
                            if (dataNewLine == i) {
                                i = 0;
                                System.out.print("\n");
                            }
                            System.out.print(value);
                            System.out.print(" ");
                            i++;
                        }

                        System.out.println("\n/////////////////////////////");
                    }
                    else {
                        //currentExcelSheet.getExcelSectionMap().remove(excelSectionItem.getHashKey());
                    }

                }
            }
            currentExcelWorkbook.excelSheetMap.put(String.valueOf(sheetIndex), currentExcelSheet);
        }
    }

    private boolean isNotInRange(int rowIndex, int columnIndex, ExcelSection excelSection)
    {
        return !(this.isInRange(rowIndex, columnIndex, excelSection));
    }

    private boolean isInRange(int rowIndex, int columnIndex, Map<String, ExcelSection> excelSectionMap)
    {
        boolean result = false;
        for (ExcelSection excelSectionItem : excelSectionMap.values()) {
            result = this.isInRange(rowIndex, columnIndex, excelSectionItem);
            if (result == true) {
                break;
            }
        }
        return result;
    }

    private boolean isNotInRange(int rowIndex, int columnIndex, Map<String, ExcelSection> excelSectionMap)
    {
        return !(this.isInRange(rowIndex, columnIndex, excelSectionMap));
    }

    private boolean isCellSpaceWithin(int rowIndex, int columnIndex, Map<String, ExcelSection> excelSectionMap)
    {
        boolean result = false;
        if (this.isInRange(rowIndex, columnIndex, excelSectionMap)) {
            result = true;
        }
        return result;
    }

    private Cell getCell(Sheet sheet, int rowIndex, int columnIndex)
    {
        Cell cell = null;
        Row row = sheet.getRow(rowIndex);
        if (row != null) {
            cell = row.getCell(columnIndex);
        }
        return cell;
    }

    private boolean isCellSpaceVertical(Sheet sheet, int rowIndex, int columnIndex)
    {
        boolean result = false;

        int rowIndexPrevious = rowIndex - 1;
        int rowIndexNext = rowIndex + 1;

        if (rowIndexPrevious <= ExcelSection.DEFAULT_VALUE && rowIndexNext <= ExcelSection.DEFAULT_VALUE) {
            Cell cellPrevious = this.getCell(sheet, rowIndexPrevious, columnIndex);
            Cell cellNext = this.getCell(sheet, rowIndexNext, columnIndex);

            if (cellPrevious != null && cellNext != null) {
                result = true;
            }
        }
        return result;
    }

    private boolean isCellSpaceHorizontal(Sheet sheet, int rowIndex, int columnIndex)
    {
        boolean result = false;

        int columnIndexPrevious = columnIndex - 1;
        int columnIndexNext = columnIndex + 1;
        if (columnIndexPrevious <= ExcelSection.DEFAULT_VALUE && columnIndexNext <= ExcelSection.DEFAULT_VALUE) {

            Cell cellPrevious = this.getCell(sheet, rowIndex, columnIndexPrevious);
            Cell cellNext = this.getCell(sheet, rowIndex, columnIndexNext);

            if (cellPrevious != null && cellNext != null) {
                result = true;
            }
        }
        return result;
    }

    private boolean isCellSpaceCross(Sheet sheet, int rowIndex, int columnIndex)
    {
        boolean result = false;
        boolean isCellSpaceVertical = this.isCellSpaceVertical(sheet, rowIndex, columnIndex);
        boolean isCellSpaceHorizontal = this.isCellSpaceHorizontal(sheet, rowIndex, columnIndex);

        if (isCellSpaceVertical && isCellSpaceHorizontal) {
            result = true;
        }
        return result;
    }

//    private int getCellSpace(Sheet sheet, Cell cell, Map<String, ExcelSection> excelSectionMap){
//        int cellSpace = ExcelSection.DEFAULT_VALUE;
//        if(this.isCellSpaceWithin(rowIndex, columnIndex, excelSectionMap)){
//            cellSpace = ExcelSection.CELL_SPACE_WITHIN;
//        }else if(this.isCellSpaceVertical(sheet, rowIndex, columnIndex)){
//            cellSpace = ExcelSection.CELL_SPACE_VERTICAL;
//        }else if(this.isCellSpaceHorizontal(sheet, rowIndex, columnIndex)){
//            cellSpace = ExcelSection.CELL_SPACE_HORIZONTAL;
//        }else if(this.isCellSpaceCross(sheet, rowIndex, columnIndex)){
//            cellSpace = ExcelSection.CELL_SPACE_CROSS;
//        }
//        return cellSpace;
//    }
    private ExcelSection getExcelSection(int rowIndex, int columnIndex, Map<String, ExcelSection> excelSectionMap, ExcelSection excelSection)
    {
        if (this.isNotInRange(rowIndex, columnIndex, excelSection)) {
            for (ExcelSection excelSectionItem : excelSectionMap.values()) {
                if (this.isInRange(rowIndex, columnIndex, excelSectionItem)) {
                    excelSection = excelSectionItem;
                    break;
                }
            }
        }

        return excelSection;
    }

    private ExcelSection getExcelSectionBasedOnCell(Cell cell, Sheet sheet, int rowIndex, int columnIndex, Map<String, ExcelSection> excelSectionMap, ExcelSection excelSection)
    {
        boolean isCellNull = cell == null;
        if (!isCellNull) {
            if (excelSection != null) {
                if (excelSection.isSet()) {
                    if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_HORIZONTAL) {
                        columnIndex = columnIndex - 1;
                    }
                    else if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_VERTICAL) {
                        rowIndex = rowIndex - 1;
                    }
                }
                excelSection = getExcelSection(rowIndex, columnIndex, excelSectionMap, excelSection);
            }
        }
        else {
            if (excelSection != null && excelSection.isSet() && excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_HORIZONTAL) {
                if (!this.isCellSpaceWithin(rowIndex, columnIndex, excelSectionMap)) {
                    this.setExcelSectionMap(excelSectionMap, excelSection);
                    excelSection = new ExcelSection();
                }
            }
            else if (excelSection != null && excelSection.isSet() && excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_VERTICAL) {
                if (this.isCellRangeNotNull(excelSection.getDataCellRange())) {
                    if (this.isCellSpaceWithin(rowIndex, columnIndex, excelSectionMap)) {
                        this.setExcelSectionMap(excelSectionMap, excelSection);
                        excelSection = new ExcelSection();
                    }
                }
            }
        }

        return excelSection;
    }

    private boolean isCellRangeNotNull(CellRangeAddress cellRangeAddress)
    {
        return (!isCellRangeNull(cellRangeAddress));
    }

    private CellRangeAddress resetCellRangeAddress(CellRangeAddress cellRangeAddress)
    {
        cellRangeAddress.setFirstColumn(ExcelSection.DEFAULT_VALUE);
        cellRangeAddress.setFirstRow(ExcelSection.DEFAULT_VALUE);
        cellRangeAddress.setLastColumn(ExcelSection.DEFAULT_VALUE);
        cellRangeAddress.setLastRow(ExcelSection.DEFAULT_VALUE);
        return cellRangeAddress;
    }

    private boolean isCellRangeNull(CellRangeAddress cellRangeAddress)
    {
        boolean result = false;
        if (cellRangeAddress.getFirstRow() == ExcelSection.DEFAULT_VALUE
                && cellRangeAddress.getFirstColumn() == ExcelSection.DEFAULT_VALUE
                && cellRangeAddress.getLastRow() == ExcelSection.DEFAULT_VALUE
                && cellRangeAddress.getLastColumn() == ExcelSection.DEFAULT_VALUE) {
            result = true;
        }
        return result;
    }

    private boolean isInRange(int rowIndex, int columnIndex, ExcelSection excelSection)
    {
        boolean result = false;
        int adjustedColumnIndex = columnIndex;
        int adjustedRowIndex = rowIndex;

        if (excelSection.isSet()) {
            if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_HORIZONTAL) {
                adjustedRowIndex = rowIndex - 1;
                if (adjustedRowIndex > excelSection.getHeaderCellRange().getFirstRow() && adjustedRowIndex < excelSection.getHeaderCellRange().getLastRow()) {
                    adjustedRowIndex = rowIndex;
                }
            }
            else if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_VERTICAL) {
                adjustedColumnIndex = columnIndex - 1;
                if (adjustedColumnIndex < excelSection.getHeaderCellRange().getFirstColumn() && adjustedColumnIndex > excelSection.getSectionCellRange().getLastColumn()) {
                    adjustedColumnIndex = columnIndex;
                }
            }
        }

        if (excelSection.getSectionCellRange().isInRange(adjustedRowIndex, adjustedColumnIndex)) {
            result = true;
        }
        return result;
    }

    private ExcelSection getExcelSectionBasedOnSheet(Sheet sheet, int sheetIndex, Workbook workbook, Map<String, ExcelSection> excelSectionMap, ExcelSection excelSection)
    {

        if (sheet != null) {
            if (!workbook.isSheetHidden(sheetIndex)) {
                if (excelSection != null && excelSection.isSet()) {
                    this.setExcelSectionMap(excelSectionMap, excelSection);
                }
                excelSection = new ExcelSection();
            }
        }

        return excelSection;
    }

    private ExcelSection getExcelSectionBasedOnRow(Row row, int sheetIndex, Map<String, ExcelSection> excelSectionMap, ExcelSection excelSection)
    {

        if (row == null) {
            if (excelSection != null && excelSection.isSet()) {
                this.setExcelSectionMap(excelSectionMap, excelSection);
                excelSection = new ExcelSection();
            }
        }

        return excelSection;
    }

    private void setCellValue(Cell cell, int sheetIndex, Map<String, String> cellValueMap)
    {
        String cellValue = cell.toString().trim();
        String key = String.valueOf(sheetIndex) + ',' + String.valueOf(cell.getRowIndex()) + ',' + String.valueOf(cell.getColumnIndex());
        if (StringUtils.isNotBlank(cellValue)) {
            cellValueMap.put(key, cellValue);
        }
    }

    private void setExcelSectionMap(Map<String, ExcelSection> excelSectionMap, ExcelSection excelSection)
    {
        if (excelSection != null) {
            if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_HORIZONTAL) {
                excelSectionMap.put(excelSection.getHashKey(), excelSection);
            }
            else if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_VERTICAL && this.isCellRangeNotNull(excelSection.getDataCellRange())) {
                excelSectionMap.put(excelSection.getHashKey(), excelSection);
            }
        }
    }

    // checks if cell value is an excel data
    private boolean isDataIdentifier(String cellValue)
    {
        boolean result = false;

        // check if numeric value
        try {
            Float.valueOf(cellValue);
            result = true;
        }
        catch (NumberFormatException ex) {
        }

        return result;
    }

    private ExcelSection setBoxCoordinates(Cell cell, int sheetIndex, ExcelSection excelSection)
    {

        if (excelSection != null && cell != null) {
            int cellRowIndex = cell.getRowIndex();
            int cellColumnIndex = cell.getColumnIndex();
            String cellValue = cell.toString().trim();

            boolean isDataIdentifier = false;
            boolean isOrientationFixed = false;
            boolean isOrientationRewindNeeded = false;

            CellRangeAddress sectionCellRange = excelSection.getSectionCellRange();
            CellRangeAddress headerCellRange = excelSection.getHeaderCellRange();
            CellRangeAddress dataCellRange = excelSection.getDataCellRange();

            int headerFirstRow = headerCellRange.getFirstRow();
            int headerLastRow = headerCellRange.getLastRow();
            int headerFirstColumn = headerCellRange.getFirstColumn();
            int headerLastColumn = headerCellRange.getLastColumn();

            int dataFirstRow = dataCellRange.getFirstRow();
            int dataLastRow = dataCellRange.getLastRow();
            int dataFirstColumn = dataCellRange.getFirstColumn();
            int dataLastColumn = dataCellRange.getLastColumn();

//            int sectionFirstRow = sectionCellRange.getFirstRow();
            int sectionLastRow = sectionCellRange.getLastRow();
//            int sectionFirstColumn = sectionCellRange.getFirstColumn();
            int sectionLastColumn = sectionCellRange.getLastColumn();

            this.setOrientation(excelSection);

            if (StringUtils.isNotBlank(cellValue)) {

                excelSection.setSet(true);

                if (headerFirstRow < 0 && headerFirstColumn < 0) {
                    headerFirstRow = cellRowIndex;
                    headerFirstColumn = cellColumnIndex;

                    headerCellRange.setFirstRow(headerFirstRow);
                    headerCellRange.setFirstColumn(headerFirstColumn);

                    sectionCellRange.setFirstRow(headerFirstRow);
                    sectionCellRange.setFirstColumn(headerFirstColumn);
                }
                else {
                    // data identifier moved...
                    isDataIdentifier = this.isDataIdentifier(cellValue);

                    if (isDataIdentifier && dataFirstRow < 0 && dataFirstColumn < 0) {

                        if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_HORIZONTAL) {
                            dataFirstRow = cellRowIndex;
                            dataFirstColumn = cellColumnIndex;

                            // orientation moved...
//                            excelSection = this.setOrientation(excelSection);

                            dataCellRange.setFirstRow(dataFirstRow);
                            dataCellRange.setFirstColumn(dataFirstColumn);

                            isOrientationFixed = true;

                            sectionCellRange.setFirstRow(headerFirstRow);
                            sectionCellRange.setFirstColumn(headerFirstColumn);
                            sectionCellRange.setLastRow(dataFirstRow);
                            sectionCellRange.setLastColumn(dataFirstColumn);
                        }
                        else if (excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_VERTICAL) {

                            dataFirstRow = cellRowIndex;
                            dataFirstColumn = cellColumnIndex;

//                            headerCellRange.setLastRow(dataFirstRow);
//                            headerCellRange.setLastColumn(dataFirstColumn);

//                            // orientation moved...
//                            excelSection = this.setOrientation(excelSection);

                            dataCellRange.setFirstRow(dataFirstRow);
                            dataCellRange.setFirstColumn(dataFirstColumn);

                            isOrientationFixed = true;
                            isOrientationRewindNeeded = true;

                            sectionCellRange.setFirstRow(headerFirstRow);
                            sectionCellRange.setFirstColumn(headerFirstColumn);
                            sectionCellRange.setLastRow(dataFirstRow);
                            sectionCellRange.setLastColumn(dataFirstColumn);
                        }

                    }
                }

                if (excelSection.isSet() && excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_HORIZONTAL) {
                    if (dataFirstRow > 0 && dataFirstColumn > 0) {
                        dataLastRow = cellRowIndex;
                        dataLastColumn = cellColumnIndex;

                        dataCellRange.setLastRow(dataLastRow);
                        dataCellRange.setLastColumn(dataLastColumn);

                        sectionCellRange.setLastRow(dataLastRow);
                        sectionCellRange.setLastColumn(dataLastColumn);
                    }
                    else {
                        if (dataFirstRow < 0 && dataFirstColumn < 0) {
                            headerLastColumn = cellColumnIndex;
                            headerCellRange.setLastColumn(headerLastColumn);
                            sectionCellRange.setLastColumn(headerLastColumn);
                            if (headerLastRow < cellRowIndex) {
                                headerLastRow = cellRowIndex;
                                headerCellRange.setLastRow(headerLastRow);
                                sectionCellRange.setLastRow(headerLastRow);
                            }
                        }
                    }
                }
                else if (excelSection.isSet() && excelSection.getHeaderOrientation() == ExcelSection.ORIENTATION_VERTICAL) {

                    if (headerFirstRow < 0 && headerFirstColumn < 0) {
                        dataLastColumn = cellColumnIndex;
                        dataCellRange.setLastColumn(dataLastColumn);
                        sectionCellRange.setLastColumn(dataLastColumn);
                        if (dataLastRow < cellRowIndex) {
                            dataLastRow = cellRowIndex;
                            dataCellRange.setLastRow(dataLastRow);
                            sectionCellRange.setLastRow(dataLastRow);
                        }

                    }
                    else {
                        if (headerFirstRow > 0 && headerFirstColumn > 0) {
                            if (dataFirstRow < 0 && dataFirstColumn < 0) {
                                headerLastRow = cellRowIndex;
                                headerLastColumn = cellColumnIndex;

                                headerCellRange.setLastRow(headerLastRow);
                                headerCellRange.setLastColumn(headerLastColumn);

                                sectionCellRange.setLastRow(headerLastRow);
                                sectionCellRange.setLastColumn(headerLastColumn);
                            }
                            else {
                                dataLastRow = cellRowIndex;
                                if (dataLastColumn < cellColumnIndex) {
                                    dataLastColumn = cellColumnIndex;
                                }

                                dataCellRange.setLastRow(dataLastRow);
                                dataCellRange.setLastColumn(dataLastColumn);

                                if (headerLastRow < dataLastRow) {
                                    headerCellRange.setLastRow(dataLastRow);
                                }

                                if (sectionLastRow < dataLastRow) {
                                    sectionCellRange.setLastRow(dataLastRow);
                                }
                                if (sectionLastColumn < dataLastColumn) {
                                    sectionCellRange.setLastColumn(dataLastColumn);
                                }
                            }
                        }
                    }
                }

            }

            excelSection.setSectionCellRange(sectionCellRange);
            excelSection.setHeaderCellRange(headerCellRange);
            excelSection.setDataCellRange(dataCellRange);

            if (isOrientationFixed) {
                this.setOrientation(excelSection);
                excelSection.setOrientationFixed(isOrientationFixed);
//                excelSection.setOrientationRewindNeeded(isOrientationRewindNeeded);
            }
        }

        return excelSection;
    }

    private ExcelSection setOrientation(ExcelSection excelSection)
    {

        CellRangeAddress headerCellRange = excelSection.getHeaderCellRange();
        CellRangeAddress dataCellRange = excelSection.getDataCellRange();

        if (excelSection != null) {
            if (!excelSection.isOrientationFixed()) {
                if (this.isCellRangeNotNull(dataCellRange)) {
                    if (headerCellRange.getFirstRow() != dataCellRange.getFirstRow()) {
                        excelSection.setHeaderOrientation(ExcelSection.ORIENTATION_HORIZONTAL);
                        excelSection.setDataOrientation(ExcelSection.ORIENTATION_VERTICAL);
                    }
                    else {
                        excelSection.setHeaderOrientation(ExcelSection.ORIENTATION_VERTICAL);
                        excelSection.setDataOrientation(ExcelSection.ORIENTATION_HORIZONTAL);
                        excelSection.setOrientationRewindNeeded(true);
                    }
                }
                else {
                    excelSection.setHeaderOrientation(ExcelSection.ORIENTATION_HORIZONTAL);
                    excelSection.setDataOrientation(ExcelSection.ORIENTATION_VERTICAL);
                }
            }
        }

        return excelSection;
    }

    private void populateMergedCellData(Sheet sheet, int sheetIndex, Map<String, String> cellValueMap, CellRangeAddress cellRangeAddress)
    {
        Cell cell = null;
        String cellValue = null;
        String key = null;
        int firstRow = cellRangeAddress.getFirstRow();
        int lastRow = cellRangeAddress.getLastRow();
        int firstColumn = cellRangeAddress.getFirstColumn();
        int lastColumn = cellRangeAddress.getLastColumn();

        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            for (int columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++) {
                key = String.valueOf(sheetIndex) + ',' + String.valueOf(rowIndex) + ',' + String.valueOf(columnIndex);

                // get cell value of the merged cell
                if (rowIndex == firstRow && columnIndex == firstColumn) {
                    cell = sheet.getRow(rowIndex).getCell(columnIndex);
                    cellValue = cell.toString().trim();
                }

                // add to cellValueMap if not null
                if (StringUtils.isNotBlank(cellValue)) {
                    cellValueMap.put(key, cellValue);
                }
                else {
                    break;
                }
            }
        }
    }

    private ExcelSection processCellValueMap(int sheetIndex, Map<String, String> cellValueMap, ExcelSection excelSection)
    {
        String key = null;
        String value = null;

        Map<String, String> headerMap = new HashMap<String, String>();
        Map<String, String> dataMap = new HashMap<String, String>();
        String mapKey = null;
        String mapValue = null;

        CellRangeAddress headerCellRange = excelSection.getHeaderCellRange();
        CellRangeAddress dataCellRange = excelSection.getDataCellRange();

        int headerFirstRow = headerCellRange.getFirstRow();
        int headerLastRow = headerCellRange.getLastRow();
        int headerFirstColumn = headerCellRange.getFirstColumn();
        int headerLastColumn = headerCellRange.getLastColumn();
        int headerOrientation = excelSection.getHeaderOrientation();

        int dataFirstRow = dataCellRange.getFirstRow();
        int dataLastRow = dataCellRange.getLastRow();
        int dataFirstColumn = dataCellRange.getFirstColumn();
        int dataLastColumn = dataCellRange.getLastColumn();
        int dataOrientation = excelSection.getDataOrientation();

        // collect headerMap from cellValueMap
        for (int rowIndex = headerFirstRow; rowIndex <= headerLastRow; rowIndex++) {
            for (int columnIndex = headerFirstColumn; columnIndex <= headerLastColumn; columnIndex++) {
                key = String.valueOf(sheetIndex) + ',' + String.valueOf(rowIndex) + ',' + String.valueOf(columnIndex);
                value = cellValueMap.get(key);

                if (StringUtils.isNotBlank(value)) {
                    mapKey = key;
                    if (excelSection.isSet()) {
                        if (headerOrientation == ExcelSection.ORIENTATION_HORIZONTAL) {
//                            mapKey = String.valueOf(sheetIndex) +","+ String.valueOf(rowIndex) +","+ headerLastColumn;
                            mapKey = String.valueOf(sheetIndex) + "," + headerLastRow + "," + String.valueOf(columnIndex);
                        }
                        else if (headerOrientation == ExcelSection.ORIENTATION_VERTICAL) {
//                            mapKey = String.valueOf(sheetIndex) +","+ headerLastRow +","+ String.valueOf(columnIndex);
                            mapKey = String.valueOf(sheetIndex) + "," + String.valueOf(rowIndex) + "," + headerLastColumn;
                        }
                    }

                    mapValue = headerMap.get(mapKey);
                    if (mapValue != null) {
                        if (!mapValue.contains(value)) {
                            mapValue = mapValue + " " + value;
                            mapValue = mapValue.trim();
                        }
                    }
                    else {
                        mapValue = value.trim();
                    }
                    headerMap.put(mapKey, mapValue);
                }
            }
        }
        excelSection.setHeaderMap(headerMap);

        // collect dataMap from cellValueMap
        for (int rowIndex = dataFirstRow; rowIndex <= dataLastRow; rowIndex++) {
            for (int columnIndex = dataFirstColumn; columnIndex <= dataLastColumn; columnIndex++) {
                key = String.valueOf(sheetIndex) + ',' + String.valueOf(rowIndex) + ',' + String.valueOf(columnIndex);
                value = cellValueMap.get(key);

                if (StringUtils.isNotBlank(value)) {

                    mapKey = key;
                    mapValue = value.trim();

//                    if(excelSection.isSet()){
//                        if(dataOrientation == ExcelSection.ORIENTATION_HORIZONTAL){
//                            mapKey = String.valueOf(sheetIndex) +","+ String.valueOf(rowIndex) +","+ columnIndex;
//                        }else if(dataOrientation == ExcelSection.ORIENTATION_VERTICAL){
//                            mapKey = String.valueOf(sheetIndex) +","+ rowIndex +","+ String.valueOf(columnIndex);
//                        }
//                    }

                    dataMap.put(mapKey, mapValue);
                }
            }
        }
        excelSection.setDataMap(dataMap);

        return excelSection;
    }
}
