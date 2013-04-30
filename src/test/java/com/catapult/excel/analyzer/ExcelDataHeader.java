package com.catapult.excel.analyzer;

import java.util.ArrayList;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class ExcelDataHeader 
{
    public static final short ORIENTATION_VERTICAL = 1;
    public static final short ORIENTATION_HORIZONTAL = 2;

    private String sheetName;
    private int sheetIndex;
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

    public ExcelDataHeader(Cell cell)
    {
        this.startRow = cell.getRowIndex();
        this.startColumn = cell.getColumnIndex();
        this.endRow = this.startRow;
        this.endColumn = this.startColumn;

        addTitle(cell.toString().trim());
    }

    public void addSubHeader(Cell cell)
    {
        // if orientation is horizontal,
        // adding a subheader is going down
        if (ORIENTATION_HORIZONTAL == orientation) {
            endRow = cell.getRowIndex();
        }
        // adding a subheader is going to the right
        else {
            endColumn = cell.getColumnIndex();
        }

        addTitle(cell.toString().trim());
    }

    @Override
    public String toString()
    {
        return new StringBuilder()
                .append("start: [").append(startRow)
                .append(",").append(startColumn).append("]")
                .append(",end: [").append(endRow)
                .append(",").append(endColumn).append("]")
                .append(",orientation: ")
                .append(",data start: [").append(dataStartRow)
                .append(",").append(dataStartColumn).append("]")
                .append(",data end : ").append(dataEndRow)
                .append(",").append(dataEndColumn).append("]")
                .append(orientation==ORIENTATION_HORIZONTAL ? "H" : "V")
                .append(",title: ").append(getTitle())
                .toString();
    }

    private void addTitle(String title)
    {
        titleList.add(title.replace("\n", " ").replaceAll("\\s{2,}", " "));
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
        return StringUtils.join(titleList, " ");
    }

    /**
     * @return the titleList
     */
    public List<String> getTitleList()
    {
        return titleList;
    }

}
