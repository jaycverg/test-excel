package com.catapult.excel.parsing;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelHeader {

    private Cell cell;

    private CellRangeAddress cellRangeAddress;

    private boolean isMergedCell;

    private int formatPoints;

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }

    public CellRangeAddress getCellRangeAddress() {
        return cellRangeAddress;
    }

    public void setCellRangeAddress(CellRangeAddress cellRangeAddress) {
        this.cellRangeAddress = cellRangeAddress;
    }

    public boolean isMergedCell() {
        return isMergedCell;
    }

    public void setMergedCell(boolean isMergedCell) {
        this.isMergedCell = isMergedCell;
    }

    public int getFormatPoints() {
        return formatPoints;
    }

    public void setFormatPoints(int formatPoints) {
        this.formatPoints = formatPoints;
    }

}
