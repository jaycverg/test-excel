package com.catapult.excel.analyzer;

import com.catapult.testexcel.CellUtil;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 *
 * @author jvergara <jvergara@gocatapult.com>
 */
public class DefaultCellAnalyzer implements CellAnalyzer
{
    private static final Pattern POSSIBLE_HEADER_TEXT = Pattern.compile("origin|destination|port|carrier|shipper|service|contract");
    private static final int SCORE_COMPARISON_OFFSET = 2; // +/- 1 offset

    private static final int HEADER_SCORE_BG = 1;
    private static final int HEADER_SCORE_MERGED = 1;
    private static final int HEADER_SCORE_BOLD_TEXT = 3;
    private static final int HEADER_SCORE_COMMON_VALUE = 5;
    private static final int HEADER_SCORE_TOTAL = HEADER_SCORE_BG + HEADER_SCORE_MERGED + HEADER_SCORE_BOLD_TEXT + HEADER_SCORE_COMMON_VALUE;

    private static final int DATA_SCORE_NUMERIC = 1;
    private static final int DATA_SCORE_COMMON_VALUE = 3;
    private static final int DATA_SCORE_TOTAL = DATA_SCORE_NUMERIC + DATA_SCORE_COMMON_VALUE;

    public void analyzeCell(CellNode cellNode)
    {
        analyzeIfData(cellNode);
        analyzeIfHeader(cellNode);
    }

    public int getHeaderMaxScore()
    {
        return HEADER_SCORE_TOTAL;
    }

    public int getDataMaxScore()
    {
        return DATA_SCORE_TOTAL;
    }

    public int getScoreComparisonOffset()
    {
        return SCORE_COMPARISON_OFFSET;
    }

    /**
     * A cell could be a header if:
     *  - it has a background
     *  - it has merged cells
     *  - it has a bold text
     *  - its value contains "origin", "destination", "port", etc.
     */
    private void analyzeIfHeader(CellNode cellNode)
    {
        if (CellUtil.hasBackground(cellNode.cell))
        {
            cellNode.headerScore += HEADER_SCORE_BG;
        }

        if (cellNode.merged)
        {
            cellNode.headerScore += HEADER_SCORE_MERGED;
        }

        if (CellUtil.isTextBold(cellNode.cell))
        {
            cellNode.headerScore += HEADER_SCORE_BOLD_TEXT;
        }

        // match for common header values
        Matcher m = POSSIBLE_HEADER_TEXT.matcher(cellNode.value.toLowerCase());
        if (m.find())
        {
            cellNode.headerScore += HEADER_SCORE_COMMON_VALUE;
        }
    }

    /**
     * A cell could be a value if:
     *  - its value is numeric
     *  - its value is an ISO country code
     */
    private void analyzeIfData(CellNode cellNode)
    {
        // try if numeric
        try
        {
            Float.parseFloat(cellNode.value);
            cellNode.dataScore += DATA_SCORE_NUMERIC;
        }
        catch(NumberFormatException ignored){}

        // match for common data values
        if (CountryCodeIndex.isCountryCode(cellNode.value))
        {
            cellNode.dataScore += DATA_SCORE_COMMON_VALUE;
        }
    }
}
