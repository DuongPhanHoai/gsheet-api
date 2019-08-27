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
}
