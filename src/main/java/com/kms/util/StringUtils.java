package com.kms.util;

public class StringUtils {
  private StringUtils() {
    throw new IllegalStateException("Utility class");
  }

  /**
   * Check if empty (also check trim is empty)
   * 
   * @param input as String
   * @return true if null or empty input string
   */
  public static boolean isEmpty(String input) {
    return input == null || input.isEmpty() || input.trim().isEmpty();
  }

  /**
   * Check if any String in Array is empty
   * 
   * @param inputs
   * @return
   */
  public static boolean isAnyEmpty(String[] inputs) {
    for (String input : inputs) {
      if (isEmpty(input))
        return true;
    }
    return false;
  }

  /**
   * Check if all strings are empty
   * 
   * @param inputs
   * @return
   */
  public static boolean isAllEmpty(String[] inputs) {
    for (String input : inputs) {
      if (!isEmpty(input))
        return false;
    }
    return true;
  }
}
