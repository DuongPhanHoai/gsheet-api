package com.kms.gdrive.sheet;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import com.kms.util.ListUtils;
import com.kms.util.StringUtils;

public class Report {
  public static final String CLASSNAME = "com.kms.gdrive.sheet.Report";
  static String TEST_NAME_COLUMN = "C";
  static String TEST_RESULT_COLUMN = "E";

  /**
   * Do not input same column for both Name and Result
   * 
   * @param TEST_NAME_COLUMN   column character (Ex: "A" or "D") default is "C"
   * @param TEST_RESULT_COLUMN column character (Ex: "A" or "D") default is "E"
   */
  static public void set_TEST_COLUMNS(String TEST_NAME_COLUMN, String TEST_RESULT_COLUMN) {
    Report.TEST_NAME_COLUMN = TEST_NAME_COLUMN;
    Report.TEST_RESULT_COLUMN = TEST_RESULT_COLUMN;
  }

  static int TEST_NAME_START_ROW = 7;

  /**
   * @param TEST_NAME_START_ROW default value is 5, min number is 2
   */
  static public void set_TEST_NAME_START_ROW(int TEST_NAME_START_ROW) {
    Report.TEST_NAME_START_ROW = TEST_NAME_START_ROW;
  }

  static final int MAX_BLANK_ROW = 5;
  static final DateTimeFormatter DATETIME_FORMATTER = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");

  // STATIC
  /**
   * find the test by name (static)
   * 
   * @param tcName              The test name to find
   * @param sheetName           The sheet to find the test
   * @param sheetID             The sheetID which can get from the google sheet
   *                            URL
   * @param allowExistingResult TRUE: find Test which has result or not; FALSE:
   *                            find the test which does not has Result
   * @return The row index of the found test ; -1 if not found
   */
  public static int findTestByName(String tcName, String sheetName, String sheetID, boolean allowExistingResult) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      return foundReport.findTestByName(tcName, sheetName, allowExistingResult);
    else
      return -1;
  }

  /**
   * update the test result by name (static)
   * 
   * @param tcName          The test name to find
   * @param tcResult        The test result to update
   * @param sheetName       The sheet to find the test
   * @param sheetID         The sheetID which can get from the google sheet URL
   * @param overWriteResult Is True, overwrite result, else the new row of test
   *                        will be created for the result
   * @return The row index of the found test and update result successful ; -1 if
   *         not found
   */
  public static int updateTestResultByName(String tcName, String tcResult, String sheetName, String sheetID,
      boolean overWriteResult) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      return foundReport.updateTestResultByName(tcName, tcResult, sheetName, overWriteResult);
    else
      return -1;
  }

  /**
   * update the test result by name (static) in the existing result column
   * 
   * @param tcName          The test name to find
   * @param tcResult        The test result to update
   * @param sheetName       The sheet to find the test
   * @param sheetID         The sheetID which can get from the google sheet URL
   * @return The row index of the found test and update result successful ; -1 if
   *         not found
   */
  public static int updateTestResultInExistingResult(String tcName, String tcResult, String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      return foundReport.updateTestResultInExistingResult(tcName, tcResult, sheetName);
    else
      return -1;
  }

  /**
   * overwrite the new test result colunm at the default location
   * (TEST_RESULT_COLUMN)
   * 
   * @param title     the header title
   * @param sheetName The sheet to find the test
   * @param sheetID   The sheetID which can get from the google sheet URL
   */
  public static void overwriteResultColHeader(String title, String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      foundReport.overwriteResultColHeader(title, sheetName);
  }

  /**
   * insert the new test result colunm at the default location
   * (TEST_RESULT_COLUMN) (static)
   * 
   * @param sheetName The sheet to find the test
   * @param sheetID   The sheetID which can get from the google sheet URL
   * @return true if the column is inserted successful
   */
  public static boolean createNewResultCol(String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      return foundReport.createNewResultCol(sheetName);
    else
      return false;
  }

  /**
   * insert the new test result colunm with title (TEST_RESULT_COLUMN) (static)
   * 
   * @param title     the header title
   * @param sheetName The sheet to find the test
   * @param sheetID   The sheetID which can get from the google sheet URL
   * @return true if the column is inserted successful
   */
  public static boolean createNewResultColTitle(String title, String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      return foundReport.createNewResultColTitle(title, sheetName);
    else
      return false;
  }

  // FACTORY
  /**
   * store the Hash of Report by the SheetID, using for factory buffer
   */
  private static HashMap<String, Report> hashReports = new HashMap<>();

  /**
   * get Report by the sheetID
   * 
   * @param sheetID The sheetID which can get from the google sheet URL
   * @return Report by the input sheetID
   */
  private static Report getReport(String sheetID) {
    if (hashReports.containsKey(sheetID))
      return hashReports.get(sheetID);
    else {
      Report newReport = new Report(sheetID);
      hashReports.put(sheetID, newReport);
      return newReport;
    }
  }

  // ****** REPORT INSTANCE ******
  /**
   * keep the sheetID after constructing
   */
  String sheetID;

  /**
   * Constructor for Report
   * 
   * @param sheetID The sheetID which can get from the sheet URL
   */
  public Report(String sheetID) {
    this.sheetID = sheetID;
  }

  /**
   * the max Row which the Test Name is not empty
   */
  int maxRowIndex = 1;

  /**
   * find the test by name
   * 
   * @param tcName              The test name to find
   * @param sheetName           The sheet to find the test
   * @param allowExistingResult TRUE: find Test which has result or not; FALSE:
   *                            find the test which does not has Result
   * @return The row index of the found test ; -1 if not found
   */
  public int findTestByName(String tcName, String sheetName, boolean allowExistingResult) {
    if (StringUtils.isAnyEmpty(new String[] { tcName, sheetName }))
      return -1;
    int blankCount = 0;
    for (int row10x = 0; row10x < 1000 && blankCount <= MAX_BLANK_ROW; row10x++) {
      // Get the Name Range
      List<List<Object>> values = Sheet.readRange(sheetName, TEST_NAME_COLUMN, TEST_NAME_START_ROW + row10x * 10,
          TEST_NAME_COLUMN, TEST_NAME_START_ROW + row10x * 10 + 10, sheetID);
      if (ListUtils.isEmpty(values))
        break; // break in blank sheet
      else {
        for (int rowIndex = 0; rowIndex < 10 && blankCount <= MAX_BLANK_ROW; rowIndex++) {
          String scanName = ListUtils.getValue(values, rowIndex, 0);
          if (!StringUtils.isEmpty(scanName)) {
            blankCount = 0;
            maxRowIndex = (TEST_NAME_START_ROW + row10x * 10 + rowIndex); // now it is current index
            if (tcName.equalsIgnoreCase(scanName) && (allowExistingResult || StringUtils.isEmpty(ListUtils.getValue(
                Sheet.readRange(sheetName, TEST_RESULT_COLUMN, maxRowIndex, TEST_RESULT_COLUMN, maxRowIndex, sheetID),
                0, 0))))
              return maxRowIndex;
          } else {
            blankCount++;
          }
        }
      }
    }
    return -1;
  }

  /**
   * update the test result by name
   * 
   * @param tcName          The test name to find
   * @param tcResult        The test result to update
   * @param sheetName       The sheet to find the test
   * @param overWriteResult Is True, overwrite result, else the new row of test
   *                        will be created for the result
   * @return The row index of the found test and update result successful ; -1 if
   *         not found
   */
  public int updateTestResultByName(String tcName, String tcResult, String sheetName, boolean overWriteResult) {
    int foundTestRow = findTestByName(tcName, sheetName, overWriteResult);
    if (foundTestRow >= 0)
      Sheet.setValue(tcResult,
          sheetName + "!" + TEST_RESULT_COLUMN + foundTestRow + ":" + TEST_RESULT_COLUMN + foundTestRow, sheetID);
    else {
      if (maxRowIndex < TEST_NAME_START_ROW)
        maxRowIndex = TEST_NAME_START_ROW;
      foundTestRow = maxRowIndex + 1;

      Sheet.setValue(tcName, sheetName + "!" + TEST_NAME_COLUMN + foundTestRow + ":" + TEST_NAME_COLUMN + foundTestRow,
          sheetID);
      Sheet.setValue(tcResult,
          sheetName + "!" + TEST_RESULT_COLUMN + foundTestRow + ":" + TEST_RESULT_COLUMN + foundTestRow, sheetID);
    }
    return foundTestRow;
  }

  /**
   * find the test by name in the existing result column
   * 
   * @param tcName    The test name to find
   * @param sheetName The sheet to find the test
   * @return The row index of the found test ; -1 if not found
   */
  public int findTestInExistingResult(String tcName, String sheetName) {
    // Get the test index from the existing result
    int existingResultIndex = testResult.findTheTestIndex(tcName);
    // Prepare to search test
    existingResultIndex++;
    int matchCount = -1;
    int blankCount = 0;
    for (int row10x = 0; row10x < 1000 && blankCount <= MAX_BLANK_ROW; row10x++) {
      // Get the Name Range
      List<List<Object>> values = Sheet.readRange(sheetName, TEST_NAME_COLUMN, TEST_NAME_START_ROW + row10x * 10,
          TEST_NAME_COLUMN, TEST_NAME_START_ROW + row10x * 10 + 10, sheetID);
      if (ListUtils.isEmpty(values))
        break; // break in blank sheet
      else {
        for (int rowIndex = 0; rowIndex < 10 && blankCount <= MAX_BLANK_ROW; rowIndex++) {
          String scanName = ListUtils.getValue(values, rowIndex, 0);
          if (!StringUtils.isEmpty(scanName)) {
            blankCount = 0;
            maxRowIndex = (TEST_NAME_START_ROW + row10x * 10 + rowIndex); // now it is current index
            if (tcName.equalsIgnoreCase(scanName)) {
              matchCount++;
              if (matchCount == existingResultIndex)
                return maxRowIndex;
            }
          } else {
            blankCount++;
          }
        }
      }
    }
    return -1;
  }

  /**
   * update the test result by name in the existing result column
   * 
   * @param tcName    The test name to find
   * @param tcResult  The test result to update
   * @param sheetName The sheet to find the test
   * @return The row index of the found test and update result successful ; -1 if
   *         not found
   */
  public int updateTestResultInExistingResult(String tcName, String tcResult, String sheetName) {
    int foundTestRow = findTestInExistingResult(tcName, sheetName);
    if (foundTestRow >= 0) {
      testResult.addNew(tcName, tcResult);
      Sheet.setValue(tcResult,
          sheetName + "!" + TEST_RESULT_COLUMN + foundTestRow + ":" + TEST_RESULT_COLUMN + foundTestRow, sheetID);
    } else {
      if (maxRowIndex < TEST_NAME_START_ROW)
        maxRowIndex = TEST_NAME_START_ROW;
      foundTestRow = maxRowIndex + 1;

      Sheet.setValue(tcName, sheetName + "!" + TEST_NAME_COLUMN + foundTestRow + ":" + TEST_NAME_COLUMN + foundTestRow,
          sheetID);
      Sheet.setValue(tcResult,
          sheetName + "!" + TEST_RESULT_COLUMN + foundTestRow + ":" + TEST_RESULT_COLUMN + foundTestRow, sheetID);
      testResult.addNew(tcName, tcResult);
    }
    return foundTestRow;
  }

  /**
   * insert the new test result colunm at the default location
   * (TEST_RESULT_COLUMN)
   * 
   * @param sheetName The sheet to find the test
   * @return true if the column is inserted successful
   */
  public boolean createNewResultCol(String sheetName) {
    // Insert a column
    Sheet.insertColumn(letterToColumn(TEST_RESULT_COLUMN), sheetName, sheetID);

    // Add the column label
    LocalDateTime now = LocalDateTime.now();
    String colHeader = now.format(DATETIME_FORMATTER);
    Sheet.setValue(colHeader, sheetName + "!" + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1) + ":"
        + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1), sheetID);
    return false;
  }

  /**
   * insert the new test result colunm at the default location
   * (TEST_RESULT_COLUMN)
   * 
   * @param title     the header title
   * @param sheetName The sheet to find the test
   * @return true if the column is inserted successful
   */
  public boolean createNewResultColTitle(String title, String sheetName) {
    // Insert a column
    Sheet.insertColumn(letterToColumn(TEST_RESULT_COLUMN), sheetName, sheetID);

    // Add the column label
    LocalDateTime now = LocalDateTime.now();
    String colHeader = title + "-" + now.format(DATETIME_FORMATTER);
    Sheet.setValue(colHeader, sheetName + "!" + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1) + ":"
        + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1), sheetID);
    return false;
  }

  /**
   * overwrite the new test result colunm at the default location
   * (TEST_RESULT_COLUMN)
   * 
   * @param title     the header title
   * @param sheetName The sheet to find the test
   */
  public void overwriteResultColHeader(String title, String sheetName) {
    // Add the column label
    LocalDateTime now = LocalDateTime.now();
    String colHeader = title + "-" + now.format(DATETIME_FORMATTER);
    Sheet.setValue(colHeader, sheetName + "!" + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1) + ":"
        + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1), sheetID);
  }

  // Follow help from:
  static public String columnToLetter(int column) {
    int temp;
    String letter = "";
    while (column > 0) {
      temp = (column) % 26;
      letter = (char) (temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }
    return letter;
  }

  static public int letterToColumn(String letter) {
    int column = 0;
    int length = letter.length();
    for (int i = 0; i < length; i++) {
      column += ((int) letter.charAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column - 1;
  }

  static class testResult {
    String name = null;
    int index = -1;

    public int getIndex() {
      return index;
    }

    String result = null;

    public testResult(String name, String result, int index) {
      this.name = name;
      this.result = result;
      this.index = index;
    }

    public boolean equal(String name) {
      if (this.name != null && name != null)
        return this.name.equalsIgnoreCase(name);
      return false;
    }

    // FACTORY
    static ArrayList<testResult> results = new ArrayList<testResult>();

    static public int findTheTestIndex(String name) {
      for (int iExistingResult = (results.size() - 1); iExistingResult >= 0; iExistingResult--)
        if (results.get(iExistingResult).equal(name))
          return results.get(iExistingResult).getIndex();
      return -1;
    }

    static public void addNew(String name, String result) {
      int foundIndex = findTheTestIndex(name) + 1;
      results.add(new testResult(name, result, foundIndex));
    }
  }
}
