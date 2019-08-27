package com.kms.gdrive.sheet;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;

import com.google.api.services.sheets.v4.model.ValueRange;
import com.kms.util.StringUtils;

public class Report {
  public static final String CLASSNAME = "com.kms.gdrive.sheet.Report";
  static final String TEST_NAME_COLUMN = "C";
  static final String TEST_RESULT_COLUMN = "E";
  static final int TEST_NAME_START_ROW = 5;
  static final int MAX_BLANK_ROW = 5;
  static final DateTimeFormatter DATETIME_FORMATTER = DateTimeFormatter.ofPattern("yyyyMMdd-HHmmss");

  // STATIC
  /**
   * find the test by name (static)
   * 
   * @param testName  The test name to find
   * @param sheetName The sheet to find the test
   * @param sheetID   The sheetID which can get from the google sheet URL
   * @return The row index of the found test ; -1 if not found
   */
  public static int findTestByName(String testName, String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      return foundReport.findTestByName(testName, sheetName);
    else
      return -1;
  }

  /**
   * update the test result by name (static)
   * 
   * @param testName  The test name to find
   * @param testName  The test result to update
   * @param sheetName The sheet to find the test
   * @param sheetID   The sheetID which can get from the google sheet URL
   * @return The row index of the found test and update result successful ; -1 if
   *         not found
   */
  public static int updateTestResultByName(String testName, String testResult, String sheetName, String sheetID) {
    Report foundReport = getReport(sheetID);
    if (foundReport != null)
      return foundReport.updateTestResultByName(testName, testResult, sheetName);
    else
      return -1;
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
  int maxRowIndex = 0;

  /**
   * find the test by name
   * 
   * @param testName  The test name to find
   * @param sheetName The sheet to find the test
   * @return The row index of the found test ; -1 if not found
   */
  public int findTestByName(String testName, String sheetName) {
    if (StringUtils.isEmpty(testName) || StringUtils.isEmpty(sheetName))
      return -1;
    int blankCount = 0;
    for (int row10x = 0; row10x < 1000 && blankCount <= MAX_BLANK_ROW; row10x++) {
      // Get the Name Range
      final String readRange = sheetName + "!" + TEST_NAME_COLUMN + (TEST_NAME_START_ROW + row10x * 10) + ":"
          + TEST_NAME_COLUMN + (TEST_NAME_START_ROW + row10x * 10 + 10);
      ValueRange testRange = Sheet.readRange(readRange, sheetID);
      List<List<Object>> values = testRange.getValues();
      if (values == null || values.isEmpty())
        break; // break in blank sheet
      else {
        for (int rowIndex = 0; rowIndex < 10 && blankCount <= MAX_BLANK_ROW; rowIndex++) {
          if (rowIndex >= values.size()) {
            blankCount = blankCount + 10 - values.size();
            break;
          }
          List<Object> row = values.get(rowIndex);
          if (row.get(0) != null) {
            String scanName = ((String) row.get(0)).trim();
            if (!scanName.isEmpty()) {
              blankCount = 0;
              maxRowIndex = (TEST_NAME_START_ROW + row10x * 10 + rowIndex);
              if (testName.equalsIgnoreCase(scanName)) {
                // Found test
                return (TEST_NAME_START_ROW + row10x * 10 + rowIndex);
              }
            } else {
              blankCount++;
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
   * update the test result by name
   * 
   * @param testName  The test name to find
   * @param testName  The test result to update
   * @param sheetName The sheet to find the test
   * @return The row index of the found test and update result successful ; -1 if
   *         not found
   */
  public int updateTestResultByName(String testName, String testResult, String sheetName) {
    int foundTestRow = findTestByName(testName, sheetName, sheetID);
    if (foundTestRow >= 0)
      Sheet.setValue(testResult,
          sheetName + "!" + TEST_RESULT_COLUMN + foundTestRow + ":" + TEST_RESULT_COLUMN + foundTestRow, sheetID);
    else {
      foundTestRow = maxRowIndex + 1;

      Sheet.setValue(testName,
          sheetName + "!" + TEST_NAME_COLUMN + foundTestRow + ":" + TEST_NAME_COLUMN + foundTestRow, sheetID);
      Sheet.setValue(testResult,
          sheetName + "!" + TEST_RESULT_COLUMN + foundTestRow + ":" + TEST_RESULT_COLUMN + foundTestRow, sheetID);
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
    Sheet.insertColumn(4, sheetID);

    // Add the column label
    LocalDateTime now = LocalDateTime.now();
    String colHeader = now.format(DATETIME_FORMATTER);
    Sheet.setValue(colHeader, sheetName + "!" + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1) + ":"
        + TEST_RESULT_COLUMN + (TEST_NAME_START_ROW - 1), sheetID);

    return false;
  }
}
