package com.catapult.excel.analyzer;

import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class ExcelDataCluster 
{
    private List<ExcelDataHeader> headers;
    private List<String> comments;
    
    private int startRow;
    private int startColumn;
    private int endRow;
    private int endColumn;

    private int dataStartRow;
    private int dataStartColumn;

    private short headerOrientation;

    ExcelDataCluster(List<ExcelDataHeader> headers)
    {
        this.headers = new ArrayList(headers);

        if (!headers.isEmpty()) {
            ExcelDataHeader first = headers.get(0);
            ExcelDataHeader last = headers.get(headers.size()-1);

            startRow = first.getStartRow();
            startColumn = first.getStartColumn();
            endRow = last.getDataEndRow();
            endColumn = last.getDataEndColumn();
            dataStartRow = first.getDataStartRow();
            dataStartColumn = first.getDataStartColumn();
        }
    }

    @Override
    public String toString()
    {
        return new StringBuilder()
                .append("cluster: ")
                .append(" [").append(startRow)
                .append(", ").append(startColumn)
                .append("] - [").append(endRow)
                .append(", ").append(endColumn)
                .append("], data: [").append(dataStartRow)
                .append(", ").append(dataStartColumn)
                .append("]")
                .toString();
    }

    /**
     * @return the headers
     */
    public List<ExcelDataHeader> getHeaders()
    {
        if (headers == null) {
            headers = new ArrayList();
        }
        return headers;
    }

    /**
     * @return the comments
     */
    public List<String> getComments()
    {
        if (comments == null) {
            comments = new ArrayList();
        }
        return comments;
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
     * @return the headerOrientation
     */
    public short getHeaderOrientation()
    {
        return headerOrientation;
    }

    /**
     * @param headerOrientation the headerOrientation to set
     */
    public void setHeaderOrientation(short headerOrientation)
    {
        this.headerOrientation = headerOrientation;
    }

}
