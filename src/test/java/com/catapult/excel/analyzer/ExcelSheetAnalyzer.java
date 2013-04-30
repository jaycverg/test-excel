package com.catapult.excel.analyzer;

import com.catapult.testexcel.CellUtil;
import java.util.ArrayList;
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
        Map<Integer, RowNode> rowMap = new HashMap();

        indexCells(rowMap, rows);
        if (rows.isEmpty()) return;

        applyMergedRegions(rowMap);
        prepareAndScoreEachCell(rows);

        CellNode firstNode = rows.get(0).first;
        rows.clear();   // clear resources
        rowMap.clear(); // clear resources
        analyzePossibleHeaders(firstNode);
    }

    private void indexCells(Map<Integer, RowNode> rowMap, List<RowNode> rows)
    {
        Iterator<Row> rowItr = sheet.rowIterator();

        while (rowItr.hasNext())
        {
            Row row = rowItr.next();
            RowNode rowNode = new RowNode();
            rowNode.index = row.getRowNum();
            Iterator<Cell> cellItr = row.cellIterator();

            while (cellItr.hasNext())
            {
                Cell cell = cellItr.next();
                if (CellUtil.getCellValue(cell) != null)
                {
                    rowNode.add(new CellNode(cell));
                }
            }

            if (!rowNode.getCellValues().isEmpty())
            {
                rows.add(rowNode);
                rowMap.put(rowNode.index, rowNode);
            }
        }
    }

    private void applyMergedRegions(Map<Integer, RowNode> rowMap)
    {
        for (int i=0; i < sheet.getNumMergedRegions(); ++i)
        {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int startRow = range.getFirstRow();
            int endRow = range.getLastRow();
            int startCol = range.getFirstColumn();
            int endCol = range.getLastColumn();

            RowNode sourceRow = rowMap.get(startRow);
            // this might be merging rows with no value
            if (sourceRow == null)
            {
                continue;
            }

            CellNode sourceCell = sourceRow.getAt(startCol);
            // this might be merging cells with no value
            if (sourceCell == null)
            {
                continue;
            }

            for (int rowIdx=startRow; rowIdx <= endRow; ++rowIdx)
            {
                RowNode rowNode = rowMap.get(rowIdx);
                for (int colIdx=startCol; colIdx <= endCol; ++colIdx)
                {
                    CellNode cn = rowNode.getAt(colIdx);
                    if (cn == null)
                    {
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
            // pack row currentNode to sort all cells
            // based on indexes and do some initializations
            currentRow.pack();
            
            currentRow.prev = prevRow;
            currentRow.next = null;
            if (prevRow != null)
            {
                prevRow.next = currentRow;
            }

            prevRow = currentRow;
            CellNode prevCell = null;
            
            for (CellNode currentCell : currentRow.getCellValues())
            {
                currentCell.parent = currentRow;
                currentCell.prev = prevCell;
                currentCell.next = null;
                if (prevCell != null)
                {
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
            if (cellNode.top != null)
            {
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
        if (style.getFillForegroundColorColor() != null)
        {
            cellNode.headerScore += 1;
        }

        if (cellNode.merged)
        {
            cellNode.headerScore += 1;
        }

        Font font = sheet.getWorkbook().getFontAt(style.getFontIndex());
        if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD)
        {
            cellNode.headerScore += 3;
        }

        Matcher m = POSSIBLE_HEADER_TEXT.matcher(cellNode.value.toLowerCase());
        if (m.find())
        {
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
        try
        {
            Float.parseFloat(cellNode.value);
            cellNode.dataScore += 1;
        }
        catch(NumberFormatException ignored){}

        if (CountryCodeIndex.isCountryCode(cellNode.value))
        {
            cellNode.dataScore += 3;
        }
    }

    private void analyzePossibleHeaders(CellNode firstNode)
    {
        List<CellNode> unprocessedList = new ArrayList();
        unprocessedList.add(firstNode);

        while(!unprocessedList.isEmpty())
        {
            CellNode cNode = unprocessedList.remove(0);
            
            if (!cNode.processed)
            {
                analyzeCellNode(cNode, unprocessedList);
            }
        }
    }

    private void analyzeCellNode(CellNode firstNode, List<CellNode> unprocessedList)
    {
        // if not a valid currentNode, just find the next blocks of data
        if (!isValidNode(firstNode))
        {
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

        // this is the usual data
        if (horizontalScore.getCount() > 1 && verticalScore.getCount() > 1)
        {
            if (horizontalScore.getHeaderAveScore() > verticalScore.getHeaderAveScore())
            {
                createHeadersHorizontal(firstNode, horizontalScore);
            }
            else
            {
                createHeadersVertical(firstNode, verticalScore);
            }
        }
        // this is for further testing processing
        else
        {
            
        }

        // find next blocks of data
        findNextDataBlockHorizontal(rightNode, unprocessedList);
        findNextDataBlockVertical(bottomNode, unprocessedList);
    }

    private boolean isValidNode(CellNode firstNode)
    {
        // a header should have at least one value either going right or down
        if (!firstNode.isNextAdjacent() && !firstNode.isNextAdjacent())
        {
            return false;
        }

        // if the bottom has an adjacent cell to the left,
        // this might not be a header
        if (firstNode.bottom != null && firstNode.bottom.isPrevAdjacent())
        {
            return false;
        }

        return true;
    }

    private CellNode computeScore(CellNode firstNode, boolean horizontal, CellNodeScore score)
    {
        CellNode lastNode = firstNode;
        
        do
        {
            lastNode.processed = true;
            score.add(lastNode);

            // applies to horizontal only
            if (horizontal && !lastNode.isNextAdjacent())
            {
                break;
            }
            // applies to vertical only
            else if (!horizontal && !lastNode.isBottomAdjacent())
            {
                break;
            }

            lastNode = horizontal ? lastNode.next : lastNode.bottom;
        }
        while (true);

        return lastNode;
    }

    /**
     * Recursively finds the next data block to right of the passed currentNode
     */
    private void findNextDataBlockHorizontal(CellNode node, List<CellNode> unprocessedList)
    {
        if (node.next != null)
        {
            if (node.isNextAdjacent())
            {
                findNextDataBlockHorizontal(node.next, unprocessedList);
            }
            else
            {
                // add to unprocessed list
                // this will be for processing
                unprocessedList.add(node.next);
            }
        }
        else if (node.bottom != null)
        {
            findNextDataBlockHorizontal(node.bottom, unprocessedList);
        }
        else
        {
            while (node.isPrevAdjacent() && node.prev.bottom == null)
            {
                node = node.prev;
            }

            if (node != null && node.bottom != null)
            {
                findNextDataBlockHorizontal(node.bottom, unprocessedList);
            }
        }
    }

    /**
     * Finds the next CellNode at the bottom of the passed currentNode
     */
    private void findNextDataBlockVertical(CellNode node, List<CellNode> unprocessedList)
    {
        if (node.parent.next != null)
        {
            unprocessedList.add(node.parent.next.first);
        }
    }

    private void createHeadersHorizontal(CellNode firstNode, CellNodeScore initialScore)
    {
        // create the headers
        Map<Integer, ExcelDataHeader> headerMap = new LinkedHashMap();
        CellNode currentNode = firstNode;
        
        do
        {
            ExcelDataHeader header = new ExcelDataHeader(currentNode.cell);
            header.setOrientation(ExcelDataHeader.ORIENTATION_HORIZONTAL);
            headerMap.put(currentNode.colIndex, header);
        }
        while(currentNode.isNextAdjacent() && (currentNode = currentNode.next) != null);

        // try to find subheaders
        if (firstNode.isBottomAdjacent())
        {
            CellNodeScore subScore = new CellNodeScore();
            currentNode = firstNode.bottom;

            do
            {
                subScore.reset();
                computeScore(currentNode, true, subScore);
                // if cells are not equal to the initial cells,
                // then its not a subheader
                if (subScore.getCount() != initialScore.getCount())
                {
                    break;
                }

                headerMap.get(currentNode.colIndex).addSubHeader(currentNode.cell);
            }
            while (currentNode.isBottomAdjacent() && (currentNode = currentNode.bottom) != null);

            // reset current node to its parent node
            currentNode = currentNode.top;
        }

        // find the first and last rows for the data of each header
        int dataStartRow = currentNode.rowIndex+1;
        int dataEndRow = dataStartRow;

        // iterate each lastHeader
        do
        {
            // determine the bottom most currentNode
            CellNode bottomNode = currentNode;

            while (bottomNode.isBottomAdjacent())
            {
                bottomNode = bottomNode.bottom;
            }

            dataEndRow = Math.max(dataEndRow, bottomNode.rowIndex);
        }
        while(currentNode.isNextAdjacent() && (currentNode = currentNode.next) != null);

        for (ExcelDataHeader header : headerMap.values())
        {
            header.setDataStartRow(dataStartRow);
            header.setDataEndRow(dataEndRow);
            header.setDataStartColumn(header.getStartColumn());
            header.setDataEndColumn(header.getStartColumn());
        }

        if (!headerMap.isEmpty())
        {
            headers.addAll(headerMap.values());
        }
    }

    private void createHeadersVertical(CellNode firstNode, CellNodeScore initialScore)
    {
        
    }

    //<editor-fold defaultstate="collapsed" desc="=== helper classes ===">
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
