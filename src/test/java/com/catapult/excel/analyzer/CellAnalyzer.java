package com.catapult.excel.analyzer;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public interface CellAnalyzer
{

    void analyzeCell(CellNode cellNode);

    int getHeaderMaxScore();

    int getDataMaxScore();

    int getScoreComparisonOffset();
}
