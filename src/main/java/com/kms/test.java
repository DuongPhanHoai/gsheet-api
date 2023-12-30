package com.kms;

import com.kms.gdrive.sheet.Report;
import com.kms.gdrive.sheet.Sheet;

public class test {
    public static void main(String[] args) {
        System.out.println("Start");
        Sheet.setCredentialDir("confDir", null);

        Report.createNewResultCol("History", "1TCSvY9aDlM4zgYuEYILJPdJzRX09l3zRTmXX70cKtWM");



    }
}
