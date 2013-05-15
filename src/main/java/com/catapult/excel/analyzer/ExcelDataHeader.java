package com.catapult.excel.analyzer;

import java.util.ArrayList;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class ExcelDataHeader implements Comparable<ExcelDataHeader>
{
    private String sheetName;
    private int sheetIndex;
    private int group;
    private int startRow;
    private int startColumn;
    private int endRow;
    private int endColumn;
    private int dataStartRow;
    private int dataStartColumn;
    private int dataEndRow;
    private int dataEndColumn;
    private short orientation;
    private List<String> titleList = new ArrayList();

    private CellNode prevCellNode;

    public ExcelDataHeader(Sheet sheet, CellNode cellNode)
    {
        this.sheetName = sheet.getSheetName();
        this.sheetIndex = sheet.getWorkbook().getSheetIndex(sheetName);
        this.startRow = cellNode.rowIndex;
        this.startColumn = cellNode.colIndex;
        this.endRow = this.startRow;
        this.endColumn = this.startColumn;

        addTitle(cellNode.cell.toString().trim());
        prevCellNode = cellNode;
    }

    void addSubHeader(CellNode cellNode)
    {
        // if orientation is horizontal,
        // adding a subheader is going down
        if (ExcelDataConstants.ORIENTATION_HORIZONTAL == orientation) {
            endRow = cellNode.rowIndex;
        }
        // adding a subheader is going to the right
        else {
            endColumn = cellNode.colIndex;
        }

        if (prevCellNode == null || prevCellNode.cell != cellNode.cell) {
            addTitle(cellNode.cell.toString().trim());
        }
        prevCellNode = cellNode;
    }

    @Override
    public String toString()
    {
        return new StringBuilder()
                .append("location: [").append(startRow)
                .append(",").append(startColumn).append("]")
                .append(" - [").append(endRow)
                .append(",").append(endColumn).append("]")
                .append(",orientation: ")
                .append(orientation==ExcelDataConstants.ORIENTATION_HORIZONTAL ? "H" : "V")
                .append(",data : [").append(dataStartRow)
                .append(",").append(dataStartColumn).append("]")
                .append(" - [").append(dataEndRow)
                .append(",").append(dataEndColumn).append("]")
                .append(",title: ").append(getTitle())
                .toString();
    }

    @Override
    public int compareTo(ExcelDataHeader o)
    {
        if (o == null) {
            return -1;
        }

        if (this.group != o.group) {
            return this.group - o.group;
        }

        if (ExcelDataConstants.ORIENTATION_HORIZONTAL == this.orientation) {
            return this.startColumn - o.startColumn;
        }
        else {
            return this.startRow - o.startRow;
        }
    }

    private void addTitle(String title)
    {
        if (!title.isEmpty()) {
            titleList.add(title.replace("\n", " ").replaceAll("\\s{2,}", " "));
        }
    }
    

    /**
     * @return the sheetName
     */
    public String getSheetName()
    {
        return sheetName;
    }

    /**
     * @param sheetName the sheetName to set
     */
    public void setSheetName(String sheetName)
    {
        this.sheetName = sheetName;
    }

    /**
     * @return the sheetIndex
     */
    public int getSheetIndex()
    {
        return sheetIndex;
    }

    /**
     * @param sheetIndex the sheetIndex to set
     */
    public void setSheetIndex(int sheetIndex)
    {
        this.sheetIndex = sheetIndex;
    }

    /**
     * @return the group
     */
    public int getGroup()
    {
        return group;
    }

    /**
     * @param group the group to set
     */
    public void setGroup(int group)
    {
        this.group = group;
    }

    /**
     * @return the startRow
     */
    public int getStartRow()
    {
        return startRow;
    }

    /**
     * @param startRow the startRow to set
     */
    public void setStartRow(int startRow)
    {
        this.startRow = startRow;
    }

    /**
     * @return the startColumn
     */
    public int getStartColumn()
    {
        return startColumn;
    }

    /**
     * @param startColumn the startColumn to set
     */
    public void setStartColumn(int startColumn)
    {
        this.startColumn = startColumn;
    }

    /**
     * @return the endRow
     */
    public int getEndRow()
    {
        return endRow;
    }

    /**
     * @param endRow the endRow to set
     */
    public void setEndRow(int endRow)
    {
        this.endRow = endRow;
    }

    /**
     * @return the endColumn
     */
    public int getEndColumn()
    {
        return endColumn;
    }

    /**
     * @param endColumn the endColumn to set
     */
    public void setEndColumn(int endColumn)
    {
        this.endColumn = endColumn;
    }

    /**
     * @return the dataStartRow
     */
    public int getDataStartRow()
    {
        return dataStartRow;
    }

    /**
     * @param dataStartRow the dataStartRow to set
     */
    public void setDataStartRow(int dataStartRow)
    {
        this.dataStartRow = dataStartRow;
    }

    /**
     * @return the dataStartColumn
     */
    public int getDataStartColumn()
    {
        return dataStartColumn;
    }

    /**
     * @param dataStartColumn the dataStartColumn to set
     */
    public void setDataStartColumn(int dataStartColumn)
    {
        this.dataStartColumn = dataStartColumn;
    }

    /**
     * @return the dataEndRow
     */
    public int getDataEndRow()
    {
        return dataEndRow;
    }

    /**
     * @param dataEndRow the dataEndRow to set
     */
    public void setDataEndRow(int dataEndRow)
    {
        this.dataEndRow = dataEndRow;
    }

    /**
     * @return the dataEndColumn
     */
    public int getDataEndColumn()
    {
        return dataEndColumn;
    }

    /**
     * @param dataEndColumn the dataEndColumn to set
     */
    public void setDataEndColumn(int dataEndColumn)
    {
        this.dataEndColumn = dataEndColumn;
    }

    /**
     * @return the orientation
     */
    public short getOrientation()
    {
        return orientation;
    }

    /**
     * @param orientation the orientation to set
     */
    public void setOrientation(short orientation)
    {
        this.orientation = orientation;
    }

    /**
     * @return the title
     */
    public String getTitle()
    {
        return StringUtils.join(titleList, " | ");
    }

    /**
     * @return the titleList
     */
    public List<String> getTitleList()
    {
        return titleList;
    }

}
