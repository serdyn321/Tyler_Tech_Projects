import java.io.*;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/*
 * Created by austin.calkins on 6/12/2017.
 */
public class CompareRecords {

    private static final String FILE_NAME = "CRM_DUPLICATES.xls";

    public static void main(String[] args) {
        try {
            File file = new File(FILE_NAME);
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(1); //selects the duplicate suspects sheet


            double previousValue = -1; //-1 is used to symbolize a non-matching record.

            //look at every duplicate record recorded
            for (int j = 1; j <= 53977; j++) {

                double percentage = CompareTwoRecords(j, sheet);

                //if the record is not a match, use the previous value
                if (percentage == -1) {
                    percentage = previousValue;
                    sheet.getRow(j).createCell(29).setCellValue(percentage);
                } else {
                    previousValue = percentage;
                    sheet.getRow(j).createCell(29).setCellValue(percentage);
                }

            }
            //writes the changes out to the file, will error if excel is open for the file
            try {
                FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
                wb.write(outputStream);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }


        } catch (Exception ioe) {
            ioe.printStackTrace();
        }
    }

    private static double CompareTwoRecords(int rowNumber, HSSFSheet sheet) {
        HSSFRow firstRow = sheet.getRow(rowNumber);
        HSSFRow secondRow = sheet.getRow(rowNumber + 1);
        double percentage = 0;
        int numElementsAdded = 0;


        //only checks for duplicates for cases 1 through 4

        if (firstRow.getCell(1).toString().equals("1.0") || firstRow.getCell(1).toString().equals("2.0") ||
                firstRow.getCell(1).toString().equals("3.0") || firstRow.getCell(1).toString().equals("4.0")) {

            // if (!firstRow.getCell(5).toString().equalsIgnoreCase("Yes")) { //if the record has not already been reviewed.

            //if the records share the same proposedMasterID
            if (firstRow.getCell(8) != null && secondRow.getCell(8) != null) {
                if (firstRow.getCell(8).toString().equals(secondRow.getCell(8).toString())) {
                    for (int col = 11; col < 29; col++) {

                        if (col == 17 || col == 28) { //ignore  job title & case Count, can add more things in this condition
                        } else {

                            //make sure there is information in both of the cells, prevents errors "NULL" is a string stored within the cell.
                            if (firstRow.getCell(col) != null && secondRow.getCell(col) != null) {

                                //these are the only operations performed after the conditions are evaluated.
                                percentage += calcPercentageSimiliarity(firstRow.getCell(col).toString(), secondRow.getCell(col).toString());
                                numElementsAdded++;
                            }
                        }

                    }
                }
            }
            //if }

        }
        //if no changes are made to calculate percentages, return a -1 to signal a no-match between records.
        if (percentage == 0 && numElementsAdded == 0) {
            return -1;
        } else {
            return percentage / (double) numElementsAdded; //returns similarity % between 0 and 1
        }
    }

    public static final double calcPercentageSimiliarity(final String s1, final String s2) {
        // during a merge, null comparisons always merge.
        if (s1.equals("NULL") || s2.equals("NULL")) {
            return 1.0;
        }
        if (s1 == null || s2 == null) {
            return 1.0;
        }

        //if they are the same, then return 100% match
            if (s1.equals(s2)) {
            return 1;
        }

        // create two work vectors of integer distances
        int[] v0 = new int[s2.length() + 1];
        int[] v1 = new int[s2.length() + 1];
        int[] vtemp;

        // initialize v0 (the previous row of distances)
        // this row is A[0][i]: edit distance for an empty s
        // the distance is just the number of characters to delete from t
        for (int i = 0; i < v0.length; i++) {
            v0[i] = i;
        }

        for (int i = 0; i < s1.length(); i++) {
            // calculate v1 (current row distances) from the previous row v0
            // first element of v1 is A[i+1][0]
            //   edit distance is delete (i+1) chars from s to match empty t
            v1[0] = i + 1;

            // use formula to fill in the rest of the row
            for (int j = 0; j < s2.length(); j++) {
                int cost = 1;
                if (s1.charAt(i) == s2.charAt(j)) {
                    cost = 0;
                }
                v1[j + 1] = Math.min(
                        v1[j] + 1,              // Cost of insertion
                        Math.min(
                                v0[j + 1] + 1,  // Cost of remove
                                v0[j] + cost)); // Cost of substitution
            }

            // copy v1 (current row) to v0 (previous row) for next iteration

            // Flip references to current and previous row
            vtemp = v0;
            v0 = v1;
            v1 = vtemp;

        }
        // now the index in v0[length of string 2] contains the number of shifts to convert the string into the other form.

        // takes the number of inserts, removes, and subs and turns it into a percentage
        if (1 - (((v0[s2.length()])) / (double) s2.length()) <= 0) {
            return 0.0;
        } else {
            return (1 - ((v0[s2.length()])) / (double) s2.length());
        }
    }
}
