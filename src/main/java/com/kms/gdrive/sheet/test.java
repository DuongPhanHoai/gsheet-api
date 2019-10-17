package com.kms.gdrive.sheet;

public class test {
  public static void main(String[] args) {
    System.out.println("Hello World!");
    Sheet.setCredentialDir("gconf", null);
    Report.updateTestResultByName("myTestcase", "Hello 123456", "TestResult", "1eaAz6HwGiZxJjpTIgasOzVegRFK2UnDUC2H6r0-SI2Q", false);
  }
}