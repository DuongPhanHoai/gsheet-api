package com;

import com.kms.gdrive.sheet.Report;

/**
 * Hello world!
 *
 */
public class App 
{
static final String sheetName = "TestResult";
static final String sheetID = "1eaAz6HwGiZxJjpTIgasOzVegRFK2UnDUC2H6r0-SI2Q";
public static void main( String[] args )
{
  // System.out.println("Hello World!");
  // Sheet.setCredentialDir("gconf", null);
  // Report.updateTestResultByName("myTestcase", "Hello from Maven" + new Date(), "TestResult", "1eaAz6HwGiZxJjpTIgasOzVegRFK2UnDUC2H6r0-SI2Q", false);
  Report.createNewResultCol(sheetName, sheetID);

}
}
