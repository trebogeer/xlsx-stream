package com.trebogeer.xlsx.stream;

/**
 * @author dimav
 *         Date: 6/22/11
 *         Time: 4:45 PM
 */
public final class XLSXConstants {

    public static final String SHEET_LOCATION = "xl/worksheets/";
    public static final String WORKBOOK_LOCATION = "xl/workbook.xml";
    public static final String WORKBOOK_REL_LOCATION = "xl/_rels/workbook.xml.rels";
    public static final String RID = "rId";

    /* Limitations */
    public static final int ROW_LIMIT = 1048576;
    public static final int COL_LIMIT = 16384;
    public static final int COL_WIDTH = 255;
    public static final int CELL_CHARACTERS_LIMIT = 32767;

    /* */
    public static final String WORKSHEET_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

    private XLSXConstants() {

    }
}
