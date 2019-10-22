package com.kms.util;

import java.util.List;

public class ListUtils {
  private ListUtils() {
    throw new IllegalStateException("Utility class");
  }

  /**
   * Check if list is empty
   * 
   * @param values values to check is empty or not
   * @return true if empty or null
   */
  public static boolean isEmpty(List<List<Object>> values) {
    return (values == null || values.isEmpty());
  }

  /**
   * Get the value from List by Row and Col indexes
   * 
   * @param values value to get the value from
   * @param rowIndex the row index
   * @param colIndex the col index
   * @return the value from List by Row and Col indexes
   */
  public static String getValue(List<List<Object>> values, int rowIndex, int colIndex) {
    if (isEmpty(values) || rowIndex >= values.size())
      return null;
    List<Object> row = values.get(rowIndex);
    if (colIndex >= row.size())
      return null;
    return ((String) row.get(colIndex)).trim();
  }
}
