package com.kms.gdrive.sheet;

import java.util.ArrayList;

class TestResult {
  String name = null;
  int index = -1;

  public int getIndex() {
    return index;
  }

  String result = null;

  public TestResult(String name, String result, int index) {
    this.name = name;
    this.result = result;
    this.index = index;
  }

  public boolean isName(String name) {
    if (this.name != null && name != null)
      return this.name.equalsIgnoreCase(name);
    return false;
  }

  // FACTORY
  static ArrayList<TestResult> results = new ArrayList<>();

  public static int findTheTestIndex(String name) {
    for (int iExistingResult = (results.size() - 1); iExistingResult >= 0; iExistingResult--)
      if (results.get(iExistingResult).isName(name))
        return results.get(iExistingResult).getIndex();
    return -1;
  }

  public static void addNew(String name, String result) {
    int foundIndex = findTheTestIndex(name) + 1;
    results.add(new TestResult(name, result, foundIndex));
  }
}