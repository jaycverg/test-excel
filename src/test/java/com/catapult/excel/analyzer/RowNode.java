package com.catapult.excel.analyzer;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
class RowNode
{
    public int index;

    public RowNode prev;
    public RowNode next;

    public CellNode first;
    public CellNode last;

    private Map<Integer,CellNode> colsMap = new HashMap();
    private List<CellNode> cols = new ArrayList();

    public void pack()
    {
        Collections.sort(cols);

        if (!cols.isEmpty())
        {
            first = cols.get(0);
            last = cols.get(cols.size()-1);
        }
    }

    public void add(CellNode value)
    {
        cols.add(value);
        colsMap.put(value.colIndex, value);
    }

    public List<CellNode> getCellValues()
    {
        return cols;
    }

    public CellNode getAt(int index)
    {
        return colsMap.get(index);
    }
}