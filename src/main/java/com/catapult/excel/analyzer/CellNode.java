package com.catapult.excel.analyzer;

import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class CellNode implements Comparable<CellNode>
{
    public int rowIndex;
    public int colIndex;

    public int headerScore;
    public int dataScore;

    public String value;
    public Cell cell;
    public boolean merged;
    public boolean header;
    public boolean processed;

    public CellNode top;
    public CellNode bottom;
    public CellNode prev;
    public CellNode next;

    public RowNode parent;

    CellNode(Cell cell)
    {
        this.rowIndex = cell.getRowIndex();
        this.colIndex = cell.getColumnIndex();
        this.value = CellUtil.getCellValueAsString(cell);
        this.cell = cell;
    }

    public boolean isPrevAdjacent()
    {
        return prev != null && prev.colIndex == this.colIndex-1;
    }

    public boolean isNextAdjacent()
    {
        return next != null && next.colIndex == this.colIndex+1;
    }

    public boolean isTopAdjacent()
    {
        return top != null && top.rowIndex == this.rowIndex-1;
    }

    public boolean isBottomAdjacent()
    {
        return bottom != null && bottom.rowIndex == this.rowIndex+1;
    }

    @Override
    public int compareTo(CellNode o)
    {
        if (o == null) {
            return -1;
        }

        return this.colIndex - o.colIndex;
    }

    @Override
    public String toString()
    {
        return new StringBuilder()
                .append("row: ").append(rowIndex)
                .append(",col: ").append(colIndex)
                .toString();
    }

}