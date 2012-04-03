package com.trebogeer.xlsx.stream;

/**
 * @author dimav
 *         Date: 6/24/11
 *         Time: 4:01 PM
 */
public final class ExcelUtils {

    private ExcelUtils() {
    }


    public static String convertNumToColString(int col) {
        // Excel counts column A as the 1st column, we
        //  treat it as the 0th one
        int excelColNum = col + 1;

        String colRef = "";
        int colRemain = excelColNum;

        while (colRemain > 0) {
            int thisPart = colRemain % 26;
            if (thisPart == 0) {
                thisPart = 26;
            }
            colRemain = (colRemain - thisPart) / 26;

            // The letter A is at 65
            char colChar = (char) (thisPart + 64);
            colRef = colChar + colRef;
        }

        return colRef;
    }

    public static String getCellReference(int col, int row) {
        return convertNumToColString(col) + (row + 1);
    }
}
