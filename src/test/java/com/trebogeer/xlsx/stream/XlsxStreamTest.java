package com.trebogeer.xlsx.stream;

import org.junit.Test;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.Calendar;
import java.util.List;
import java.util.Random;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * @author dimav
 *         Date: 2/10/12
 *         Time: 3:42 PM
 */
public class XlsxStreamTest {

    @Test
    public void test() throws Exception {

        long start = System.currentTimeMillis();

        String tmpDir = System.getProperty("java.io.tmpdir");
        String outXlsxName = "template-out.xlsx";
        OutputStream os = new BufferedOutputStream(new FileOutputStream(tmpDir + System.getProperty("file.separator") + outXlsxName));
        ZipOutputStream zos = new ZipOutputStream(os);
        ZipFile templateZip = new ZipFile(getClass().getResource("/template.xlsx").getFile());
        try {
            WorkbookWriter wb = WorkbookWriter.getWorkbookWriter(templateZip, zos);
            List<? extends ColumnHeader> headers = Arrays.asList(Headers.values());
            SpreadSheetWriter sw = wb.addNextSheet("Sheet 1", false, headers);
            generateDataSheet(sw);
            wb.flush();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            zos.finish();
            zos.close();
            templateZip.close();
        }
        System.out.println("Time : "+ (System.currentTimeMillis() - start));
    }


    private void generateDataSheet(SpreadSheetWriter sw) throws Exception {

        Random rnd = new Random();
        Calendar calendar = Calendar.getInstance();
        sw.beginSheet();

        //write data rows
        for (int rownum = 1; rownum < 100000/*0*/; rownum++) {
            sw.insertNextRow();

            sw.createCell(0, "Hello, " + rownum + "!");
            sw.createCell(1, (double) rnd.nextInt(100) / 100);
            sw.createCell(2, (double) rnd.nextInt(10) / 10);
            sw.createCell(3, rnd.nextInt(10000));
            sw.createCell(4, calendar.getTimeInMillis());

            sw.endRow();

            calendar.roll(Calendar.DAY_OF_YEAR, 1);
        }
        sw.endSheet();
    }

    private enum Headers implements ColumnHeader {
        TITLE("Header 1"),
        CHANGE("Header 2"),
        RATIO("One More Header"),
        EXPENSES("One More Header 2"),
        DATE("And The Last Header");

        String name;

        Headers(String title) {
            name = title;
        }

        public String getName() {
            return name;
        }
    }
}
