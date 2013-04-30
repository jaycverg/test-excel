package com.catapult.excel.analyzer;

import com.catapult.testexcel.CellUtil;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class ExcelSheetAnalyzer
{
    private static final Pattern POSSIBLE_HEADER_TEXT = Pattern.compile("origin|destination|port|carrier|shipper|service|contract");
    private static final int SCORE_COMPARISON_OFFSET = 1; // +/- 1 offset

    private Sheet sheet;
    private List<ExcelDataHeader> headers = new ArrayList();

    public ExcelSheetAnalyzer(Sheet sheet)
    {
        this.sheet = sheet;
    }

    public List<ExcelDataHeader> getHeaders()
    {
        return headers;
    }

    public void analyze()
    {
        headers.clear();

        List<RowNode> rows = new ArrayList();
        Map<Integer, RowNode> rowsMap = new HashMap();

        indexCells(rowsMap, rows);
        if (rows.isEmpty()) return;

        applyMergedRegions(rowsMap);
        prepareAndScoreEachCell(rows);
        analyzePossibleHeaders(rowsMap, rows);
    }

    private void indexCells(Map<Integer, RowNode> rowsMap, List<RowNode> rows)
    {
        Iterator<Row> rowItr = sheet.rowIterator();

        while (rowItr.hasNext())
        {
            Row row = rowItr.next();
            RowNode rowNode = new RowNode();
            rowNode.index = row.getRowNum();
            Iterator<Cell> cellItr = row.cellIterator();

            while (cellItr.hasNext()) {
                Cell cell = cellItr.next();
                if (CellUtil.getCellValue(cell) != null) {
                    rowNode.add(new CellNode(cell));
                }
            }

            if (!rowNode.getCellValues().isEmpty()) {
                rows.add(rowNode);
                rowsMap.put(rowNode.index, rowNode);
            }
        }
    }

    private void applyMergedRegions(Map<Integer, RowNode> rowsMap)
    {
        for (int i=0; i < sheet.getNumMergedRegions(); ++i)
        {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int startRow = range.getFirstRow();
            int endRow = range.getLastRow();
            int startCol = range.getFirstColumn();
            int endCol = range.getLastColumn();

            RowNode sourceRow = rowsMap.get(startRow);
            // this might be merging rows with no value
            if (sourceRow == null) {
                continue;
            }

            CellNode sourceCell = sourceRow.getAt(startCol);
            // this might be merging cells with no value
            if (sourceCell == null) {
                continue;
            }

            for (int rowIdx=startRow; rowIdx <= endRow; ++rowIdx)
            {
                RowNode rowNode = rowsMap.get(rowIdx);
                for (int colIdx=startCol; colIdx <= endCol; ++colIdx)
                {
                    CellNode cn = rowNode.getAt(colIdx);
                    if (cn == null) {
                        cn = new CellNode(sourceCell.cell);
                        cn.rowIndex = rowIdx;
                        cn.colIndex = colIdx;
                        rowNode.add(cn);
                    }
                    cn.merged = true;
                }
            }
        }
    }

    /**
     * Analyzes each cell and set its linked nodes
     */
    private void prepareAndScoreEachCell(List<RowNode> rows)
    {
        RowNode prevRow = null;
        for (RowNode currentRow : rows)
        {
            // pack row node to sort all cells
            // based on indexes and do some initializations
            currentRow.pack();
            
            currentRow.prev = prevRow;
            currentRow.next = null;
            if (prevRow != null) {
                prevRow.next = currentRow;
            }

            prevRow = currentRow;
            CellNode prevCell = null;
            
            for (CellNode currentCell : currentRow.getCellValues())
            {
                currentCell.parent = currentRow;
                currentCell.prev = prevCell;
                currentCell.next = null;
                if (prevCell != null) {
                    prevCell.next = currentCell;
                }

                prevCell = currentCell;
                findTopCellNode(currentCell);
                analyzeIfData(currentCell);
                analyzeIfHeader(currentCell);
            }
        }
    }

    private void findTopCellNode(CellNode cellNode)
    {
        cellNode.top = null;
        cellNode.bottom = null;
        RowNode topRow = cellNode.parent.prev;
        
        while (topRow != null)
        {
            cellNode.top = topRow.getAt(cellNode.colIndex);
            if (cellNode.top != null) {
                cellNode.top.bottom = cellNode;
                break;
            }
            topRow = topRow.prev;
        }
    }

    /**
     * A cell could be a header if:
     *  - (+1) it has a background
     *  - (+1) it has merged cells
     *  - (+3) it has a bold text
     *  - (+5) its value contains "origin", "destination", "port", etc.
     */
    private void analyzeIfHeader(CellNode cellNode)
    {
        Cell cell = cellNode.cell;
        CellStyle style = cell.getCellStyle();
        if (style.getFillForegroundColorColor() != null) {
            cellNode.headerScore += 1;
        }

        if (cellNode.merged) {
            cellNode.headerScore += 1;
        }

        Font font = sheet.getWorkbook().getFontAt(style.getFontIndex());
        if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD) {
            cellNode.headerScore += 3;
        }

        Matcher m = POSSIBLE_HEADER_TEXT.matcher(cellNode.value.toLowerCase());
        if (m.find()) {
            cellNode.headerScore += 5;
        }
    }

    /**
     * A cell could be a value if:
     *  - (+1) its value is numeric
     *  - (+3) its value is an ISO country code
     */
    private void analyzeIfData(CellNode cellNode)
    {
        // try if numeric
        try {
            Float.parseFloat(cellNode.value);
            cellNode.dataScore += 1;
        }
        catch(NumberFormatException ignored){}

        if (CountryCodeIndex.isCountryCode(cellNode.value)) {
            cellNode.dataScore += 3;
        }
    }

    private void analyzePossibleHeaders(Map<Integer, RowNode> rowsMap, List<RowNode> rows)
    {
        List<CellNode> unprocessedList = new ArrayList();
        unprocessedList.add(rows.get(0).first);

        while(!unprocessedList.isEmpty())
        {
            CellNode cNode = unprocessedList.remove(0);
            
            if (!cNode.processed) {
                analyzeCellNode(cNode, unprocessedList);
            }
        }
    }

    private void analyzeCellNode(CellNode firstNode, List<CellNode> unprocessedList)
    {
        // - if no adjacent cells, just find the next blocks of data
        // - for this algorithm, a header should have at least one
        //   value either going right or down
        if (!firstNode.isNextAdjacent() && !firstNode.isNextAdjacent()) {
            firstNode.processed = true;
            findNextDataBlockHorizontal(firstNode, unprocessedList);
            findNextDataBlockVertical(firstNode, unprocessedList);
            return;
        }

        CellNode rightNode;
        CellNode bottomNode;

        // find horizontally the possible headers
        CellNodeScore horizontalScore = new CellNodeScore();
        rightNode = computeScore(firstNode, true, horizontalScore);

        // find vertically the possible headers
        CellNodeScore verticalScore = new CellNodeScore();
        bottomNode = computeScore(firstNode, false, verticalScore);

        if (horizontalScore.getCount() > 1 && verticalScore.getCount() > 1) {
            if (horizontalScore.getHeaderAveScore() > verticalScore.getHeaderAveScore()) {
                createHeadersHorizontal(firstNode, horizontalScore);
            }
            else {
                createHeadersVertical(firstNode, verticalScore);
            }
        }
        // this is for further testing processing
        else {
            
        }

        // find next blocks of data
        findNextDataBlockHorizontal(rightNode, unprocessedList);
        findNextDataBlockVertical(bottomNode, unprocessedList);
    }

    private CellNode computeScore(CellNode firstNode, boolean horizontal, CellNodeScore score)
    {
        CellNode lastNode = firstNode;
        
        do
        {
            lastNode.processed = true;
            score.add(lastNode);

            // applies to horizontal only
            if (horizontal && !lastNode.isNextAdjacent()) {
                break;
            }
            // applies to vertical only
            else if (!horizontal && !lastNode.isBottomAdjacent()) {
                break;
            }

            lastNode = horizontal ? lastNode.next : lastNode.bottom;
        }
        while (true);

        return lastNode;
    }

    /**
     * Recursively finds the next data block to right of the passed node
     */
    private void findNextDataBlockHorizontal(CellNode node, List<CellNode> unprocessedList)
    {
        if (node.next != null) {
            if (node.isNextAdjacent()) {
                findNextDataBlockHorizontal(node.next, unprocessedList);
            }
            else {
                // add to unprocessed list
                // this will be for processing
                unprocessedList.add(node.next);
            }
        }
        else if (node.bottom != null) {
            findNextDataBlockHorizontal(node.bottom, unprocessedList);
        }
        else {
            while (node.isPrevAdjacent() && node.prev.bottom == null)
            {
                node = node.prev;
            }

            if (node != null && node.bottom != null) {
                findNextDataBlockHorizontal(node.bottom, unprocessedList);
            }
        }
    }

    /**
     * Finds the next CellNode at the bottom of the passed node
     */
    private void findNextDataBlockVertical(CellNode node, List<CellNode> unprocessedList)
    {
        if (node.parent.next != null) {
            unprocessedList.add(node.parent.next.first);
        }
    }

    private void createHeadersHorizontal(CellNode firstNode, CellNodeScore initialScore)
    {
        // create the headers
        Map<Integer, ExcelDataHeader> headersMap = new LinkedHashMap();
        CellNode node = firstNode;
        
        do
        {
            ExcelDataHeader header = new ExcelDataHeader(node.cell);
            headersMap.put(node.colIndex, header);
        }
        while(node.isNextAdjacent() && (node = node.next) != null);

        // try to find subheaders
        CellNodeScore subScore = new CellNodeScore();
        node = firstNode.bottom;
        do
        {
            subScore.reset();
            computeScore(node, true, subScore);
            // if cells are not equal to the initial cells,
            // then its not a subheader
            if (subScore.getCount() != initialScore.getCount()) {
                break;
            }

            headersMap.get(node.colIndex).addSubHeader(node.cell);
        }
        while (node.isBottomAdjacent() && (node = node.bottom) != null);
    }

    private void createHeadersVertical(CellNode firstNode, CellNodeScore initialScore)
    {
        
    }

    //<editor-fold defaultstate="collapsed" desc="=== helper classes ===">
    private static class RowNode
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

    private static class CellNode implements Comparable<CellNode>
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

    private static class CellNodeScore
    {
        private int count;
        private int headerScore;
        private int dataScore;

        public void reset()
        {
            count = 0;
            headerScore = 0;
            dataScore = 0;
        }

        public void add(CellNode cn)
        {
            count++;
            headerScore += cn.headerScore;
            dataScore += cn.dataScore;
        }

        public double getHeaderAveScore()
        {
            return headerScore / count;
        }

        public double getDataAveScore()
        {
            return dataScore / count;
        }

        public int getCount()
        {
            return count;
        }

    }
    //</editor-fold>
}
