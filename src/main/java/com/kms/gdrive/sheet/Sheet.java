package com.kms.gdrive.sheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.InsertDimensionRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.ValueRange;
import com.kms.util.StringUtils;

public class Sheet {
  public static final String CLASSNAME = "com.kms.gdrive.sheet.Sheet";
  private static final String APPLICATION_NAME = "KMS Google Sheet API";
  private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
  private static String CREDENTIALS_DIRECTORY = "gconf";
  private static String CREDENTIALS_FILE = "gsheet-auth.json"; // As resource
  // private static String credentialsFilePath = "gsheet-auth.json";
  private static final List<String> SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS, SheetsScopes.DRIVE);

  public static void setCredentialDir (String CREDENTIALS_DIRECTORY, String CREDENTIALS_FILE) {
    if (!StringUtils.isEmpty(CREDENTIALS_DIRECTORY))
      Sheet.CREDENTIALS_DIRECTORY = CREDENTIALS_DIRECTORY;
    if (!StringUtils.isEmpty(CREDENTIALS_FILE))
      Sheet.CREDENTIALS_FILE = CREDENTIALS_FILE;
  }

  /**
   * Creates an authorized Credential object.
   * 
   * @param httpTransport The network HTTP Transport.
   * @return An authorized Credential object.
   */
  private static Credential getCredentials(final NetHttpTransport httpTransport) {
    try {
      // Disable log

      // Load client secrets
      InputStreamReader credentialReader = null;
      final java.util.logging.Logger buggyLogger = java.util.logging.Logger
          .getLogger(FileDataStoreFactory.class.getName());
      buggyLogger.setLevel(java.util.logging.Level.SEVERE);

      // Load google report configuration directory path / env
      File checkExists = new File(CREDENTIALS_DIRECTORY + File.separator + CREDENTIALS_FILE);

      boolean resourceMode = true;
      if (checkExists.exists() && checkExists.isFile())
        resourceMode = false;
      
      credentialReader = resourceMode?
        new InputStreamReader(Sheet.class.getResourceAsStream(File.separator+CREDENTIALS_FILE)):
        new InputStreamReader((new FileInputStream(CREDENTIALS_DIRECTORY + File.separator + CREDENTIALS_FILE)));
      File tokenDirFileObj = resourceMode?
        new File(Sheet.class.getResource(File.separator).getFile()):
        new File(CREDENTIALS_DIRECTORY);

      GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, credentialReader);

      // Build flow and trigger user authorization request.
      GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(httpTransport, JSON_FACTORY,
          clientSecrets, SCOPES).setDataStoreFactory(new FileDataStoreFactory(tokenDirFileObj))
              .setAccessType("offline").build();
      LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
      return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
    } catch (Exception e) {
      Logger.getLogger(CLASSNAME).log(Level.WARNING, e.getMessage());
      return null;
    }
  }

  /**
   * the the range by sheetID
   * 
   * @param sheetName The sheet to find the test
   * @param startCol  start Column to get range
   * @param startRow  start Row to get range
   * @param endCol    end Column to get range
   * @param endRow    end Row to get range
   * @param sheetID   The sheetID which can get from the google sheet URL
   * @return the range
   */
  public static List<List<Object>> readRange(String sheetName, String startCol, int startRow, String endCol, int endRow,
      String sheetID) {
    Sheet foundSheet = getSheet(sheetID);
    if (foundSheet != null)
      return foundSheet.readRange(sheetName, startCol, startRow, endCol, endRow);
    else
      return Collections.emptyList();
  }

  /**
   * setValue get the values from range (static)
   * 
   * @param value      To write to as String
   * @param writeRange The range to write
   * @param sheetID    The sheetID which can get from the google sheet URL
   * @return true is successful
   */
  public static boolean setValue(String value, String writeRange, String sheetID) {
    Sheet foundSheet = getSheet(sheetID);
    if (foundSheet != null)
      return foundSheet.setValue(value, writeRange);
    else
      return false;
  }

  /**
   * insert a column at the index (static)
   * 
   * @param columnIndex The index of the column to insert
   * @param sheetID     The sheetID which can get from the google sheet URL
   * @return true is successful
   */
  public static boolean insertColumn(int columnIndex, String sheetID) {
    Sheet foundSheet = getSheet(sheetID);
    if (foundSheet != null)
      return foundSheet.insertColumn(columnIndex);
    else
      return false;
  }

  // MANAGE Sheet object by Factory
  /**
   * store the Hash of Sheet by the SheetID, using for factory buffer
   */
  private static HashMap<String, Sheet> hashSheets = new HashMap<>();

  /**
   * get Sheet by the sheetID
   * 
   * @param sheetID The sheetID which can get from the google sheet URL
   * @return Sheet by the input sheetID
   */
  private static Sheet getSheet(String sheetID) {
    if (hashSheets.containsKey(sheetID))
      return hashSheets.get(sheetID);
    else {
      Sheet newSheet = new Sheet(sheetID);
      hashSheets.put(sheetID, newSheet);
      return newSheet;
    }
  }

  // OBJECT declaration
  /**
   * has to be constructed by constructor for service as Sheets
   */
  Sheets service = null;

  /**
   * keep the sheetID after constructing
   */
  String sheetID = "";

  /**
   * Constructor for Sheet
   * 
   * @param sheetID The sheetID which can get from the sheet URL
   */
  public Sheet(String sheetID) {
    try {
      this.sheetID = sheetID;
      NetHttpTransport.Builder transportBuilder = new NetHttpTransport.Builder();
      NetHttpTransport httpTransport = transportBuilder.build();
      transportBuilder.doNotValidateCertificate();
      service = new Sheets.Builder(httpTransport, JSON_FACTORY, getCredentials(httpTransport))
          .setApplicationName(APPLICATION_NAME).build();
    } catch (Exception e) {
      Logger.getLogger(CLASSNAME).log(Level.WARNING, e.getMessage());
    }
  }

  /**
   * read the range from the input
   * 
   * @param sheetName The sheet to find the test
   * @param startCol  start Column to get range
   * @param startRow  start Row to get range
   * @param endCol    end Column to get range
   * @param endRow    end Row to get range
   * @return the range
   */
  public List<List<Object>> readRange(String sheetName, String startCol, int startRow, String endCol, int endRow) {
    if (service != null)
      try {
        final String readRange = sheetName + "!" + startCol + startRow + ":" + endCol + endRow;
        ValueRange valueRange = service.spreadsheets().values().get(sheetID, readRange).execute();
        return valueRange.getValues();
      } catch (IOException e) {
        Logger.getLogger(CLASSNAME).log(Level.WARNING, e.getMessage());
      }
    return Collections.emptyList();
  }

  /**
   * setValue set the values to range
   * 
   * @param value      To write to as String
   * @param writeRange The range to write
   * @return true is successful
   */
  public boolean setValue(String value, String writeRange) {
    if (service != null)
      try {
        // Create value list range
        ValueRange updateValues = new ValueRange();
        updateValues.setValues(Arrays.asList(Arrays.asList((Object) value)));
        service.spreadsheets().values().update(sheetID, writeRange, updateValues).setValueInputOption("USER_ENTERED")
            .execute();
        return true;
      } catch (IOException e) {
        Logger.getLogger(CLASSNAME).log(Level.WARNING, e.getMessage());
      }
    return false;
  }

  /**
   * insert a column at the index
   * 
   * @param columnIndex The index of the column to insert
   * @return true is successful
   */
  public boolean insertColumn(int columnIndex) {
    Spreadsheet spreadsheet = null;
    if (service != null)
      // Get sheet id
      try {
        spreadsheet = service.spreadsheets().get(sheetID).execute();

        Integer isheetID = spreadsheet.getSheets().get(0).getProperties().getSheetId();

        // Set column insert
        DimensionRange dimentionRange = new DimensionRange();
        dimentionRange.setStartIndex(columnIndex);
        dimentionRange.setEndIndex(columnIndex + 1);
        dimentionRange.setSheetId(isheetID);
        dimentionRange.setDimension("COLUMNS");

        InsertDimensionRequest insertCol = new InsertDimensionRequest();
        insertCol.setRange(dimentionRange);

        // Execute to insert column
        BatchUpdateSpreadsheetRequest r = new BatchUpdateSpreadsheetRequest()
            .setRequests(Arrays.asList(new Request().setInsertDimension(insertCol)));
        service.spreadsheets().batchUpdate(sheetID, r).execute();
        return true;
      } catch (IOException e) {
        Logger.getLogger(CLASSNAME).log(Level.WARNING, e.getMessage());
      }
    return false;
  }
}
