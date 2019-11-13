package com.kms.gdrive.sheet;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;

import com.kms.util.ListUtils;
import com.kms.util.StringUtils;

public class Report {
  public static final String CLASSNAME = "com.kms.gdrive.sheet.Report";
  static String testNameCol = "C";
  static String testResultCol = "E";

  /**
   * Do not input same column for both Name and Result
   * 
   * @param testNameCol   column character (Ex: "A" or "D") default is "C"
   * @param testResultCol column character (Ex: "A" or "D") default is "E"
   */
  public static void setTestCols(String testNameCol, String testResultCol) {
    Report.testNameCol = testNameCol;
    Report.testResultCol = testResultCol;
  }

  static int testNameStartRow = 12;

  /**
   * @param testNameStartRow default value is 12, min number is 2
   */
  public static void setTestNameStartRow(int testNameStartRow) {
    Report.testNameStartRow = testNameStartRow;
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
    return (foundReport == null)?-1:foundReport.findTestByName(tcName, sheetName, allowExistingResult);
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
    return (foundReport == null)?-1:foundReport.updateTestResultByName(tcName, tcResult, sheetName, overWriteResult);
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
    return (foundReport == null)?-1:foundReport.updateTestResultInExistingResult(tcName, tcResult, sheetName);
  }

  /**
   * overwrite the new test result colunm at the default location
   * (testResultCol)
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
   * (testResultCol) (static)
   * 
   * @param sheetName The sheet to find the test
   * @param sheetID   The sheetID which can get from the google sheet URL
   * @return true if the column is inserted successful
   */
  public static boolean createNewResultCol(String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    return (foundReport != null)&&foundReport.createNewResultCol(sheetName);
  }

  /**
   * insert the new test result colunm with title (testResultCol) (static)
   * 
   * @param title     the header title
   * @param sheetName The sheet to find the test
   * @param sheetID   The sheetID which can get from the google sheet URL
   * @return true if the column is inserted successful
   */
  public static boolean createNewResultColTitle(String title, String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    return (foundReport != null)&&foundReport.createNewResultColTitle(title, sheetName);
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
      List<List<Object>> values = Sheet.readRange(sheetName, testNameCol, testNameStartRow + row10x * 10,
          testNameCol, testNameStartRow + row10x * 10 + 10, sheetID);
      if (ListUtils.isEmpty(values))
        break; // break in blank sheet
      else {
        for (int rowIndex = 0; rowIndex < 10 && blankCount <= MAX_BLANK_ROW; rowIndex++) {
          String scanName = ListUtils.getValue(values, rowIndex, 0);
          if (!StringUtils.isEmpty(scanName)) {
            blankCount = 0;
            maxRowIndex = (testNameStartRow + row10x * 10 + rowIndex); // now it is current index
            if (tcName.equalsIgnoreCase(scanName) && (allowExistingResult || StringUtils.isEmpty(ListUtils.getValue(
                Sheet.readRange(sheetName, testResultCol, maxRowIndex, testResultCol, maxRowIndex, sheetID),
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
    return updateTestResultAtRow(tcName, tcResult, sheetName, findTestByName(tcName, sheetName, overWriteResult));
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
    int existingResultIndex = TestResult.findTheTestIndex(tcName);
    // Prepare to search test
    existingResultIndex++;
    int matchCount = -1;
    int blankCount = 0;
    for (int row10x = 0; row10x < 1000 && blankCount <= MAX_BLANK_ROW; row10x++) {
      // Get the Name Range
      List<List<Object>> values = Sheet.readRange(sheetName, testNameCol, testNameStartRow + row10x * 10,
          testNameCol, testNameStartRow + row10x * 10 + 10, sheetID);
      if (ListUtils.isEmpty(values))
        break; // break in blank sheet
      else {
        for (int rowIndex = 0; rowIndex < 10 && blankCount <= MAX_BLANK_ROW; rowIndex++) {
          String scanName = ListUtils.getValue(values, rowIndex, 0);
          if (!StringUtils.isEmpty(scanName)) {
            blankCount = 0;
            maxRowIndex = (testNameStartRow + row10x * 10 + rowIndex); // now it is current index
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
    return updateTestResultAtRow(tcName, tcResult, sheetName, findTestInExistingResult(tcName, sheetName));
  }

  public int updateTestResultAtRow(String tcName, String tcResult, String sheetName, int rowIndex) {
    int foundTestRow = rowIndex;
    if (foundTestRow >= 0) {
      TestResult.addNew(tcName, tcResult);
      Sheet.setValue(tcResult,
          sheetName + "!" + testResultCol + foundTestRow + ":" + testResultCol + foundTestRow, sheetID);
    } else {
      if (maxRowIndex < testNameStartRow)
        maxRowIndex = testNameStartRow;
      foundTestRow = maxRowIndex + 1;

      Sheet.setValue(tcName, sheetName + "!" + testNameCol + foundTestRow + ":" + testNameCol + foundTestRow,
          sheetID);
      Sheet.setValue(tcResult,
          sheetName + "!" + testResultCol + foundTestRow + ":" + testResultCol + foundTestRow, sheetID);
      TestResult.addNew(tcName, tcResult);
    }
    return foundTestRow;
  }

  /**
   * insert the new test result colunm at the default location
   * (testResultCol)
   * 
   * @param sheetName The sheet to find the test
   * @return true if the column is inserted successful
   */
  public boolean createNewResultCol(String sheetName) {
    // Record the current formula of the header
    List<List<Object>> values = null;
    if (testNameStartRow > 2)
      values = Sheet.readRange(sheetName, testResultCol, 1, testResultCol, testNameStartRow-2, sheetID);

    // Insert a column
    Sheet.insertColumn(letterToColumn(testResultCol), sheetName, sheetID);
    

    // Write down the old formula
    if (values != null && testNameStartRow > 2)
      Sheet.setValues(values, sheetName, testResultCol, 1, testResultCol, testNameStartRow-2, sheetID);

    // Add the column label
    LocalDateTime now = LocalDateTime.now();
    String colHeader = now.format(DATETIME_FORMATTER);
    Sheet.setValue(colHeader, sheetName + "!" + testResultCol + (testNameStartRow - 1) + ":"
        + testResultCol + (testNameStartRow - 1), sheetID);
    return false;
  }

  /**
   * insert the new test result colunm at the default location
   * (testResultCol)
   * 
   * @param title     the header title
   * @param sheetName The sheet to find the test
   * @return true if the column is inserted successful
   */
  public boolean createNewResultColTitle(String title, String sheetName) {
    // Insert a column
    Sheet.insertColumn(letterToColumn(testResultCol), sheetName, sheetID);

    // Add the column label
    LocalDateTime now = LocalDateTime.now();
    String colHeader = title + "-" + now.format(DATETIME_FORMATTER);
    Sheet.setValue(colHeader, sheetName + "!" + testResultCol + (testNameStartRow - 1) + ":"
        + testResultCol + (testNameStartRow - 1), sheetID);
    return false;
  }

  /**
   * overwrite the new test result colunm at the default location
   * (testResultCol)
   * 
   * @param title     the header title
   * @param sheetName The sheet to find the test
   */
  public void overwriteResultColHeader(String title, String sheetName) {
    // Add the column label
    LocalDateTime now = LocalDateTime.now();
    String colHeader = title + "-" + now.format(DATETIME_FORMATTER);
    Sheet.setValue(colHeader, sheetName + "!" + testResultCol + (testNameStartRow - 1) + ":"
        + testResultCol + (testNameStartRow - 1), sheetID);
  }

  // Follow help from:
  public static String columnToLetter(int column) {
    int temp;
    StringBuilder letter = new StringBuilder();
    while (column > 0) {
      temp = (column) % 26;
      letter.insert(0, (char) (temp + 65));
      column = (column - temp - 1) / 26;
    }
    return letter.toString();
  }

  public static int letterToColumn(String letter) {
    int column = 0;
    double length = letter.length();
    for (int i = 0; i < length; i++) {
      column += ((int) letter.charAt(i) - 64) * Math.pow(26, length - i - 1);
    }
    return column - 1;
  }
}
