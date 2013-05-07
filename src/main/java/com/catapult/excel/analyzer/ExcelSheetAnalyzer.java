package com.catapult.excel.analyzer;

import com.catapult.testexcel.CellUtil;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class ExcelSheetAnalyzer
{
    private Sheet sheet;
    private List<ExcelDataHeader> headers = new ArrayList();

    private CellAnalyzer cellAnalyzer;

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

        CellNode firstNode = rows.get(0).firstChild;
        rows.clear();   // clear resources
        rowMap.clear(); // clear resources
        analyzePossibleHeaders(firstNode);

        Collections.sort(headers);
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
                getCellAnalyzer().analyzeCell(currentCell);
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
        firstNode.processed = true;

        CellNode rightNode = firstNode;
        CellNode bottomNode = firstNode;

        processNode:
        {
            if (!isValidNode(firstNode))
            {
                break processNode;
            }

            // check style consistency
            boolean hStyleConsistent = checkHorizontalStyleConsistency(firstNode);
            boolean vStyleConsistent = checkVerticalStyleConsistency(firstNode);

            if (!hStyleConsistent && !vStyleConsistent)
            {
                break processNode;
            }

            // find horizontally the possible headers
            CellNodeScore horizontalScore = new CellNodeScore();
            rightNode = computeScoreHorizontal(firstNode, horizontalScore);

            // this might be a header comment
            if (horizontalScore.count == 1 && !vStyleConsistent) {
                break processNode;
            }

            // find vertically the possible headers
            CellNodeScore verticalScore = new CellNodeScore();
            bottomNode = computeScoreVertical(firstNode, verticalScore);

            if (horizontalScore.getHeaderScorePercentage() > verticalScore.getHeaderScorePercentage()) {
                createHeadersHorizontal(firstNode, verticalScore);
                markHeaderGroupHorizontalAsProcessed(firstNode);
            }
            // vertical orientation
            else {
                createHeadersVertical(firstNode, verticalScore);
                markHeaderGroupVerticalAsProcessed(firstNode);
            }
        }

        // find next blocks of data for processing
        findNextDataBlockHorizontal(rightNode, unprocessedList);
        findNextDataBlockVertical(bottomNode, unprocessedList);
    }

    private boolean isValidNode(CellNode firstNode)
    {
        // a header should have at least one value either going right or down
        if (!firstNode.isNextAdjacent() && !firstNode.isBottomAdjacent())
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

    private CellNode computeScoreHorizontal(CellNode firstNode, CellNodeScore score)
    {
        CellNode lastNode = firstNode;

        do
        {
            score.add(lastNode);
        }
        while (lastNode.isNextAdjacent() && (lastNode = lastNode.next) != null);

        return lastNode;
    }

    private CellNode computeScoreVertical(CellNode firstNode, CellNodeScore score)
    {
        CellNode lastNode = firstNode;

        do
        {
            score.add(lastNode);
        }
        while (lastNode.isBottomAdjacent() && (lastNode = lastNode.bottom) != null);

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
            unprocessedList.add(node.parent.next.firstChild);
        }
    }

    private void createHeadersHorizontal(CellNode firstNode, CellNodeScore initialScore)
    {
        // create the headers
        Map<Integer, ExcelDataHeader> headerMap = new LinkedHashMap();
        CellNode currentNode = firstNode;

        // iterate rightward
        do
        {
            ExcelDataHeader header = new ExcelDataHeader(currentNode);
            header.setOrientation(ExcelDataHeader.ORIENTATION_HORIZONTAL);
            headerMap.put(currentNode.colIndex, header);
        }
        while(currentNode.isNextAdjacent() && (currentNode = currentNode.next) != null);

        // try to find subheaders
        if (firstNode.isBottomAdjacent())
        {
            CellNodeScore subScore = new CellNodeScore();
            currentNode = firstNode.bottom;
            double initialHeaderAve = initialScore.getHeaderAveScore();

            do
            {
                subScore.reset();
                computeScoreHorizontal(currentNode, subScore);
                // if cells are not equal to the initial cells,
                // then its not a subheader
                if (subScore.count != initialScore.count)
                {
                    break;
                }

                // if header average score does not match
                if (!scoreMatches(initialHeaderAve, subScore.getHeaderAveScore()))
                {
                    break;
                }

                // add subheaders
                CellNode rightNode = currentNode;
                do
                {
                    headerMap.get(rightNode.colIndex).addSubHeader(rightNode);
                }
                while(rightNode.isNextAdjacent() && (rightNode = rightNode.next) != null);
            }
            while (currentNode.isBottomAdjacent() && (currentNode = currentNode.bottom) != null);

            // reset current node to its parent node
            currentNode = currentNode.top;
        }

        // find the firstChild and lastChild rows for the data of each header
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
            headerMap.clear();
        }
    }

    private void createHeadersVertical(CellNode firstNode, CellNodeScore initialScore)
    {
        // create the headers
        Map<Integer, ExcelDataHeader> headerMap = new LinkedHashMap();
        CellNode currentNode = firstNode;

        // iterate downward
        do
        {
            ExcelDataHeader header = new ExcelDataHeader(currentNode);
            header.setOrientation(ExcelDataHeader.ORIENTATION_VERTICAL);
            headerMap.put(currentNode.rowIndex, header);
        }
        while(currentNode.isBottomAdjacent() && (currentNode = currentNode.bottom) != null);

        // try to find subheaders
        if (firstNode.isNextAdjacent())
        {
            CellNodeScore subScore = new CellNodeScore();
            currentNode = firstNode.next;
            double initialHeaderAve = initialScore.getHeaderAveScore();

            do
            {
                subScore.reset();
                computeScoreVertical(currentNode, subScore);
                // if cells are not equal to the initial cells,
                // then its not a subheader
                if (subScore.count != initialScore.count)
                {
                    break;
                }

                // if header average score does not match
                if (!scoreMatches(initialHeaderAve, subScore.getHeaderAveScore()))
                {
                    break;
                }

                // add subheaders
                CellNode bottomNode = currentNode;
                do
                {
                    headerMap.get(bottomNode.rowIndex).addSubHeader(bottomNode);
                }
                while(bottomNode.isNextAdjacent() && (bottomNode = bottomNode.bottom) != null);
            }
            while (currentNode.isNextAdjacent() && (currentNode = currentNode.next) != null);

            // reset current node to its parent node
            currentNode = currentNode.prev;
        }

        // find the firstChild and lastChild rows for the data of each header
        int dataStartCol = currentNode.colIndex+1;
        int dataEndCol = dataStartCol;

        // iterate each lastHeader
        do
        {
            // determine the right most currentNode
            CellNode rightNode = currentNode;

            while (rightNode.isNextAdjacent())
            {
                rightNode = rightNode.next;
            }

            dataEndCol = Math.max(dataEndCol, rightNode.colIndex);
        }
        while(currentNode.isBottomAdjacent() && (currentNode = currentNode.bottom) != null);

        for (ExcelDataHeader header : headerMap.values())
        {
            header.setDataStartColumn(dataStartCol);
            header.setDataEndColumn(dataEndCol);
            header.setDataStartRow(header.getStartRow());
            header.setDataEndRow(header.getEndRow());
        }

        if (!headerMap.isEmpty())
        {
            headers.addAll(headerMap.values());
            headerMap.clear();
        }
    }

    private boolean scoreMatches(double source, double target)
    {
        int offset = getCellAnalyzer().getScoreComparisonOffset();
        return (target >= source - offset && target <= source + offset);
    }

    private boolean checkHorizontalStyleConsistency(CellNode firstNode)
    {
        boolean isTextBold = CellUtil.isTextBold(firstNode.cell);
        boolean hasBackground = CellUtil.hasBackground(firstNode.cell);

        while (firstNode.isNextAdjacent() && (firstNode = firstNode.next) != null)
        {
            if (isTextBold != CellUtil.isTextBold(firstNode.cell)) {
                return false;
            }
            if (hasBackground != CellUtil.hasBackground(firstNode.cell)) {
                return false;
            }
        }

        return true;
    }

    private boolean checkVerticalStyleConsistency(CellNode firstNode)
    {
        boolean isTextBold = CellUtil.isTextBold(firstNode.cell);
        boolean hasBackground = CellUtil.hasBackground(firstNode.cell);

        while (firstNode.isBottomAdjacent() && (firstNode = firstNode.bottom) != null)
        {
            if (isTextBold != CellUtil.isTextBold(firstNode.cell)) {
                return false;
            }
            if (hasBackground != CellUtil.hasBackground(firstNode.cell)) {
                return false;
            }
        }

        return true;
    }

    private void markHeaderGroupHorizontalAsProcessed(CellNode firstNode)
    {
        do
        {
            firstNode.processed = true;
            
            CellNode bottom = firstNode;
            while (bottom.isBottomAdjacent())
            {
                bottom = bottom.bottom;
                bottom.processed = true;
            }
        }
        while (firstNode.isNextAdjacent() && (firstNode = firstNode.next) != null);
    }

    private void markHeaderGroupVerticalAsProcessed(CellNode firstNode)
    {
        do
        {
            firstNode.processed = true;

            CellNode right = firstNode;
            while (right.isNextAdjacent())
            {
                right = right.next;
                right.processed = true;
            }
        }
        while (firstNode.isBottomAdjacent() && (firstNode = firstNode.bottom) != null);
    }

    /**
     * @return the cellAnalyzer
     */
    public CellAnalyzer getCellAnalyzer()
    {
        if (cellAnalyzer == null) {
            cellAnalyzer = new DefaultCellAnalyzer();
        }
        return cellAnalyzer;
    }

    /**
     * @param cellAnalyzer the cellAnalyzer to set
     */
    public void setCellAnalyzer(CellAnalyzer cellAnalyzer)
    {
        this.cellAnalyzer = cellAnalyzer;
    }

    //<editor-fold defaultstate="collapsed" desc="=== helper classes ===">
    private class CellNodeScore
    {
        int count;
        int headerScore;
        int dataScore;

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

        public double getHeaderScorePercentage()
        {
            int maxScore = getCellAnalyzer().getHeaderMaxScore();
            return ((headerScore / (double)count) / (double)maxScore) * 100;
        }

        public double getDataAveScore()
        {
            return dataScore / count;
        }

        public double getDataScorePercentage()
        {
            int maxScore = getCellAnalyzer().getDataMaxScore();
            return ((dataScore / (double)count) / (double)maxScore) * 100;
        }

    }
    //</editor-fold>
}
