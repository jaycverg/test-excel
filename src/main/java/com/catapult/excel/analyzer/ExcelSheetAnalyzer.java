package com.catapult.excel.analyzer;

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

    private int headerGroupCounter;

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
        headerGroupCounter = 1;
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

    /**
     * Indexes all not empty cells
     */
    private void indexCells(Map<Integer, RowNode> rowMap, List<RowNode> rows)
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
                RowNode rowNode = rowMap.get(rowIdx);
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
            // pack row currentNode to sort all cells
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
                getCellAnalyzer().analyzeCell(currentCell);
            }
        }
    }

    private void findTopCellNode(CellNode cellNode)
    {
        cellNode.top = null;
        cellNode.bottom = null;
        RowNode topRow = cellNode.parent.prev;
        
        while (topRow != null) {
            cellNode.top = topRow.getAt(cellNode.colIndex);
            if (cellNode.top != null) {
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

        while (!unprocessedList.isEmpty()) {
            CellNode cNode = unprocessedList.remove(0);

            if (!cNode.processed) {
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
            if (!isValidNode(firstNode)) {
                break processNode;
            }

            // check style consistency
            boolean hStyleConsistent = checkHorizontalStyleConsistency(firstNode);
            boolean vStyleConsistent = checkVerticalStyleConsistency(firstNode);

            // if no consistent styles, this might not be a header
            if (!hStyleConsistent && !vStyleConsistent) {
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

            boolean isHorizontal = horizontalScore.getHeaderScorePercentage() > verticalScore.getHeaderScorePercentage();

            // verify if it is not really horizontal
            if (!isHorizontal) {
                if (horizontalScore.complexHeader && !verticalScore.complexHeader) {
                    isHorizontal = true;
                }
                else if (hStyleConsistent && vStyleConsistent && horizontalScore.count > verticalScore.count)
                {
                    isHorizontal = true;
                }
            }

            if (isHorizontal) {
                if (createHeadersHorizontal(firstNode, rightNode, horizontalScore)) {
                    markHeaderGroupHorizontalAsProcessed(firstNode);
                }
                else {
                    rightNode = firstNode;
                    bottomNode = firstNode;
                }
            }
            // vertical orientation
            else {
                if (createHeadersVertical(firstNode, bottomNode, verticalScore))  {
                    markHeaderGroupVerticalAsProcessed(firstNode);
                }
                else {
                    rightNode = firstNode;
                    bottomNode = firstNode;
                }
            }
        }

        // find next blocks of data for processing
        if (rightNode != firstNode) {
            findNextDataBlockHorizontal(rightNode, unprocessedList);
        }
        findNextDataBlockVertical(bottomNode, unprocessedList);
    }

    private boolean isValidNode(CellNode firstNode)
    {
        // a header should have at least one value either going rightNode or down
        if (!firstNode.isNextAdjacent() && !firstNode.isBottomAdjacent())
        {
            return false;
        }

        // if the bottomNode has an adjacent cell to the left,
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

        // find possibe more nodes to right
        if (lastNode.isBottomAdjacent()) 
        {
            int count = 0;
            int bottomRowIndex = score.complexHeaderLastIndex;
            CellNode right;
            CellNode bottom = lastNode.bottom;
            
            findMoreHeaders:
            do {
                while (bottom.isBottomAdjacent() && !bottom.isNextAdjacent())
                {
                    bottom = bottom.bottom;
                }

                // break if right node is not adjacent or null
                if (!bottom.isNextAdjacent()) break;

                right = bottom.next;
                bottomRowIndex = Math.max(bottomRowIndex, bottom.rowIndex);

                do {
                    count++;
                }
                while (!right.isTopAdjacent() && right.isNextAdjacent() && (right = right.next) != null);

                if (right.isTopAdjacent()) {
                    CellNode top = right.top;

                    do {
                        if (top.rowIndex == lastNode.rowIndex) {
                            score.count += count - 1;
                            lastNode = computeScoreHorizontal(top, score);
                            score.complexHeader = true;
                            score.complexHeaderLastIndex = bottomRowIndex;
                            break findMoreHeaders;
                        }
                    }
                    while (top.isTopAdjacent() && (top = top.top) != null);
                }

                bottom = right;
            }
            while (true); // end do-while loop
        }

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
     * Recursively finds the next data block to the right of the passed node
     */
    private void findNextDataBlockHorizontal(CellNode node, List<CellNode> unprocessedList)
    {
        if (node.next != null)
        {
            if (node.isNextAdjacent()) {
                findNextDataBlockHorizontal(node.next, unprocessedList);
            }
            else {
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
            while (node.isPrevAdjacent() && node.prev.bottom == null) {
                node = node.prev;
            }

            if (node != null && node.bottom != null) {
                findNextDataBlockHorizontal(node.bottom, unprocessedList);
            }
        }
    }

    /**
     * Finds the next data block at the bottom of the passed node
     */
    private void findNextDataBlockVertical(CellNode node, List<CellNode> unprocessedList)
    {
        if (node.parent.next != null) {
            unprocessedList.add(node.parent.next.firstChild);
        }
    }

    private boolean createHeadersHorizontal(CellNode firstNode, CellNode lastNode, CellNodeScore initialScore)
    {
        // create the headers
        Map<Integer, ExcelDataHeader> headerMap = new LinkedHashMap();

        // keep the currentNode
        // this value will change as we navigate the nodes
        CellNode currentNode = firstNode;

        // iterate rightward
        CellNode rightNode = currentNode;
        do
        {
            ExcelDataHeader header = new ExcelDataHeader(sheet, rightNode);
            header.setOrientation(ExcelDataHeader.ORIENTATION_HORIZONTAL);
            headerMap.put(rightNode.colIndex, header);
        }
        while(rightNode.colIndex < lastNode.colIndex && (rightNode = rightNode.next) != null);

        // try to find subheaders
        if (currentNode.isBottomAdjacent())
        {
            // determine the max merged row index that where merged down from headers
            int maxMergedRowIndex = Math.max(currentNode.rowIndex, initialScore.complexHeaderLastIndex);

            // iterate rightward
            rightNode = currentNode;
            do {
                if (rightNode.merged) {
                    // check if bottomNode node is merged with the top
                    CellNode bottomNode = rightNode.bottom;
                    while (bottomNode != null && bottomNode.merged && bottomNode.cell == bottomNode.top.cell) {
                        maxMergedRowIndex = Math.max(maxMergedRowIndex, bottomNode.rowIndex);
                        bottomNode = bottomNode.bottom;
                    }
                }
            }
            while (rightNode.colIndex < lastNode.colIndex && (rightNode = rightNode.next) != null);

            // apply subheaders that are merged with the header
            if (maxMergedRowIndex > currentNode.rowIndex)
            {
                CellNode bottomNode = currentNode.bottom;
                while (bottomNode.rowIndex <= maxMergedRowIndex)
                {
                    // iterate rightward
                    rightNode = bottomNode;
                    do {
                        ExcelDataHeader header = headerMap.get(rightNode.colIndex);
                        if (header != null) {
                            header.addSubHeader(rightNode);
                        }
                        else {
                            header = new ExcelDataHeader(sheet, rightNode);
                            header.setOrientation(ExcelDataHeader.ORIENTATION_HORIZONTAL);
                            headerMap.put(rightNode.colIndex, header);
                        }
                    }
                    while (rightNode.colIndex < lastNode.colIndex && (rightNode = rightNode.next) != null);

                    bottomNode = bottomNode.bottom;
                }

                // reset to top node
                currentNode = bottomNode.top;
            }

            // find more subheaders that might qualify as headers
            CellNodeScore subScore = new CellNodeScore();
            double initialHeaderAve = initialScore.getHeaderAveScore();

            while (currentNode.isBottomAdjacent() && (currentNode = currentNode.bottom) != null)
            {
                checkBottomNode:
                {
                    subScore.reset();
                    computeScoreHorizontal(currentNode, subScore);

                    // if cells are greater than the initial cells,
                    // the initial cells were just header comments
                    if (subScore.count > initialScore.count) {
                        return false;
                    }

                    // if cells are not equal to the initial cells,
                    // then its not a subheader
                    if (subScore.count != initialScore.count) {
                        break checkBottomNode;
                    }

                    // if header average score does not match
                    if (!scoreMatches(initialHeaderAve, subScore.getHeaderAveScore())) {
                        break checkBottomNode;
                    }

                    // add subheaders
                    rightNode = currentNode;
                    do {
                        ExcelDataHeader header = headerMap.get(rightNode.colIndex);
                        if (header != null) {
                            header.addSubHeader(rightNode);
                        }
                        else {
                            header = new ExcelDataHeader(sheet, rightNode);
                            header.setOrientation(ExcelDataHeader.ORIENTATION_HORIZONTAL);
                            headerMap.put(rightNode.colIndex, header);
                        }
                    }
                    while(rightNode.colIndex < lastNode.colIndex && (rightNode = rightNode.next) != null);

                    // continue while loop
                    continue;

                }  // end labeled block

                // reset current node to its parent node
                currentNode = currentNode.top;
                break;
                
            } // end while loop
        } // end if

        // find the firstChild and lastChild rows for the data of each header
        int dataStartRow = currentNode.rowIndex+1;
        int dataEndRow = dataStartRow;

        // iterate each lastHeader
        rightNode = currentNode;
        do {
            // determine the bottom most currentNode
            CellNode bottomNode = rightNode;

            while (bottomNode.isBottomAdjacent()) {
                bottomNode = bottomNode.bottom;
            }

            dataEndRow = Math.max(dataEndRow, bottomNode.rowIndex);
        }
        while (rightNode.colIndex < lastNode.colIndex && (rightNode = rightNode.next) != null);

        for (ExcelDataHeader header : headerMap.values()) {
            header.setDataStartRow(dataStartRow);
            header.setDataEndRow(dataEndRow);
            header.setDataStartColumn(header.getStartColumn());
            header.setDataEndColumn(header.getStartColumn());
            header.setGroup(headerGroupCounter);
        }

        if (!headerMap.isEmpty()) {
            headers.addAll(headerMap.values());
            headerMap.clear();

            // increment counter if there are headers
            headerGroupCounter++;
        }

        return true;
    }

    private boolean createHeadersVertical(CellNode firstNode, CellNode lastNode, CellNodeScore initialScore)
    {
        // create the headers
        Map<Integer, ExcelDataHeader> headerMap = new LinkedHashMap();
        CellNode currentNode = firstNode;

        // iterate downward
        CellNode bottomNode = currentNode;
        do {
            ExcelDataHeader header = new ExcelDataHeader(sheet, bottomNode);
            header.setOrientation(ExcelDataHeader.ORIENTATION_VERTICAL);
            headerMap.put(bottomNode.rowIndex, header);
        }
        while (bottomNode.rowIndex < lastNode.rowIndex && (bottomNode = bottomNode.bottom) != null);

        // try to find subheaders
        if (currentNode.isNextAdjacent())
        {
            CellNodeScore subScore = new CellNodeScore();
            double initialHeaderAve = initialScore.getHeaderAveScore();

            while (currentNode.isNextAdjacent() && (currentNode = currentNode.next) != null)
            {
                checkRightNode:
                {
                    subScore.reset();
                    computeScoreVertical(currentNode, subScore);
                    // if cells are not equal to the initial cells,
                    // then its not a subheader
                    if (subScore.count != initialScore.count)
                    {
                        break checkRightNode;
                    }

                    // if header average score does not match
                    if (!scoreMatches(initialHeaderAve, subScore.getHeaderAveScore()))
                    {
                        break checkRightNode;
                    }

                    // add subheaders
                    bottomNode = currentNode;
                    do
                    {
                        ExcelDataHeader header = headerMap.get(bottomNode.rowIndex);
                        if (header != null) {
                            header.addSubHeader(bottomNode);
                        }
                    }
                    while(bottomNode.rowIndex < lastNode.rowIndex && (bottomNode = bottomNode.bottom) != null);

                    // continue while loop
                    continue;

                } // end labeled block

                // reset current node to its parent node
                currentNode = currentNode.prev;
                break;
                
            } // end while loop

        } // end if

        // find the firstChild and lastChild rows for the data of each header
        int dataStartCol = currentNode.colIndex+1;
        int dataEndCol = dataStartCol;

        // iterate each lastHeader
        bottomNode = currentNode;
        do {
            // determine the rightNode most currentNode
            CellNode rightNode = bottomNode;

            while (rightNode.isNextAdjacent()) {
                rightNode = rightNode.next;
            }

            dataEndCol = Math.max(dataEndCol, rightNode.colIndex);
        }
        while (bottomNode.rowIndex < lastNode.rowIndex && (bottomNode = bottomNode.bottom) != null);

        for (ExcelDataHeader header : headerMap.values()) {
            header.setDataStartColumn(dataStartCol);
            header.setDataEndColumn(dataEndCol);
            header.setDataStartRow(header.getStartRow());
            header.setDataEndRow(header.getEndRow());
            header.setGroup(headerGroupCounter);
        }

        if (!headerMap.isEmpty()) {
            headers.addAll(headerMap.values());
            headerMap.clear();

            // increment counter if there are headers
            headerGroupCounter++;
        }

        return true;
    }

    private boolean scoreMatches(double source, double target)
    {
        int offset = getCellAnalyzer().getScoreComparisonOffset();
        return (target >= source - offset);
    }

    private boolean checkHorizontalStyleConsistency(CellNode firstNode)
    {
        boolean isTextBoldConsistent = true;
        boolean isBackgroundConsistent = true;
        
        boolean isTextBold = CellUtil.isTextBold(firstNode.cell);
        boolean hasBackground = CellUtil.hasBackground(firstNode.cell);

        while (firstNode.isNextAdjacent() && (firstNode = firstNode.next) != null) {
            isTextBoldConsistent  = (isTextBold == CellUtil.isTextBold(firstNode.cell));
            isBackgroundConsistent = (hasBackground == CellUtil.hasBackground(firstNode.cell));
        }

        return isTextBoldConsistent || isBackgroundConsistent;
    }

    private boolean checkVerticalStyleConsistency(CellNode firstNode)
    {
        boolean isTextBoldConsistent = true;
        boolean isBackgroundConsistent = true;

        boolean isTextBold = CellUtil.isTextBold(firstNode.cell);
        boolean hasBackground = CellUtil.hasBackground(firstNode.cell);

        while (firstNode.isBottomAdjacent() && (firstNode = firstNode.bottom) != null) {
            isTextBoldConsistent  = (isTextBold == CellUtil.isTextBold(firstNode.cell));
            isBackgroundConsistent = (hasBackground == CellUtil.hasBackground(firstNode.cell));
        }

        return isTextBoldConsistent || isBackgroundConsistent;
    }

    private void markHeaderGroupHorizontalAsProcessed(CellNode firstNode)
    {
        do {
            firstNode.processed = true;

            CellNode bottom = firstNode;
            while (bottom.isBottomAdjacent()) {
                bottom = bottom.bottom;
                bottom.processed = true;
            }
        }
        while (firstNode.isNextAdjacent() && (firstNode = firstNode.next) != null);
    }

    private void markHeaderGroupVerticalAsProcessed(CellNode firstNode)
    {
        do {
            firstNode.processed = true;

            CellNode right = firstNode;
            while (right.isNextAdjacent()) {
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
        int complexHeaderLastIndex;
        boolean complexHeader;

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
