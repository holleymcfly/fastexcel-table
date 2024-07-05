package de.herrmann.holger.fastexcel;

import org.dhatim.fastexcel.Workbook;
import org.dhatim.fastexcel.Worksheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.LinkedList;

public class TableTest {

    public final static String OUTPUT_FILE_WORKING = "C:/temp/fastexcel-table-working.xlsx";
    public final static String OUTPUT_FILE_CORRUPT = "C:/temp/fastexcel-table-corrupt.xlsx";

    private Workbook workbook;
    private Worksheet worksheet;

    /**
     * Create two files <pre>OUTPUT_FILE_WORKING</pre> and <pre>OUTPUT_FILE_CORRUPT.</pre>
     * The first one will have a formatted table with everything working as expected.
     * The second one will have an ampersand in one of the headers, which leads to an error message when opening the excel file.
     * Both will have an ampersand in its data table, which seems to be no problem.
     */
    public static void main(String[] args) throws IOException {
        new TableTest().run();
    }

    private void run() throws IOException {

        // Create the working Excel file.
        try (FileOutputStream fos = new FileOutputStream(OUTPUT_FILE_WORKING)) {
            createWorkbook(fos);
            int row = addDataToWorkbook();
            worksheet.range(0, 0, row - 1, getDataRow1().size() - 1).createTable(toArray(getHeadersWorking()))
                    .styleInfo().setStyleName("TableStyleMedium3");
            workbook.finish();
        }

        // Create the corrupt Excel file. Does hardly the same as above, but only uses other headers and another output path.
        try (FileOutputStream fos = new FileOutputStream(OUTPUT_FILE_CORRUPT)) {
            createWorkbook(fos);
            int row = addDataToWorkbook();
            worksheet.range(0, 0, row - 1, getDataRow1().size() - 1).createTable(toArray(getHeadersCorrupt()))
                    .styleInfo().setStyleName("TableStyleMedium3");
            workbook.finish();
        }
    }

    private int addDataToWorkbook() {

        int row = 0;
        row = addRowToWorksheet(worksheet, getDataRow1(), row);
        row = addRowToWorksheet(worksheet, getDataRow2(), row);
        row = addRowToWorksheet(worksheet, getDataRow3(), row);
        return row;
    }

    private void createWorkbook(FileOutputStream fos) {

        workbook = new Workbook(fos, "FastExcel-Table", "1.0");
        worksheet = workbook.newWorksheet("FastExcel-Table");
    }

    private int addRowToWorksheet(Worksheet worksheet, LinkedList<String> entries, int row) {

        int col = 0;
        for (String data : entries) {
            worksheet.value(row, col, data);
            col++;
        }
        row++;
        return row;
    }

    private LinkedList<String> getHeadersWorking() {
        return new LinkedList<>(Arrays.asList("A column", "Another column", "A column without ampersand"));
    }

    private LinkedList<String> getHeadersCorrupt() {
        return new LinkedList<>(Arrays.asList("A column", "Another column", "A column & an ampersand"));
    }

    private LinkedList<String> getDataRow1() {
        return new LinkedList<>(Arrays.asList("Row 1 & Column 1", "Row 1 & Column 2", "Row 1 & Column 3"));
    }

    private LinkedList<String> getDataRow2() {
        return new LinkedList<>(Arrays.asList("Row 2 & Column 1", "Row 2 & Column 2", "Row 2 & Column 3"));
    }

    private LinkedList<String> getDataRow3() {
        return new LinkedList<>(Arrays.asList("Row 3 & Column 1", "Row 3 & Column 2", "Row 3 & Column 3"));
    }

    private String[] toArray(LinkedList<String> list) {

        Object[] objectArray = list.toArray();
        int length = objectArray.length;
        String[] stringArray = new String[length];
        for (int i = 0; i < length; i++) {
            stringArray[i] = (String) objectArray[i];
        }

        return stringArray;
    }
}
