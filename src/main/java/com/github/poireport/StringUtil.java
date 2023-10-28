package com.github.poireport;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringUtil {

    private static final Map<Character, Character> UPPERCASE_MAP = new HashMap<>();

    static {
        UPPERCASE_MAP.put(';', '；');
        UPPERCASE_MAP.put(':', '：');
        UPPERCASE_MAP.put(')', '）');
        UPPERCASE_MAP.put('?', '？');
        UPPERCASE_MAP.put('!', '！');
        UPPERCASE_MAP.put('"', '“');
        UPPERCASE_MAP.put('(', '（');
        UPPERCASE_MAP.put(',', '，');
        UPPERCASE_MAP.put('[', '【');
        UPPERCASE_MAP.put(']', '】');
        UPPERCASE_MAP.put('.', '。');
    }

    private static final Pattern LINE_PATTERN = Pattern.compile("[_-](\\w)");
    private static final Pattern HUMP_PATTERN = Pattern.compile("[A-Z]");

    public static Integer[] parseNumber(String str) {
        if (str == null) {
            return new Integer[0];
        }
        List<Integer> result = new ArrayList<>();
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < str.length(); i++) {
            char c = str.charAt(i);
            if (c >= '0' && c <= '9') {
                builder.append(c);
            } else if (builder.length() > 0) {
                result.add(Integer.valueOf(builder.toString()));
                builder.setLength(0);
            }
        }
        if (builder.length() > 0) {
            result.add(Integer.valueOf(builder.toString()));
        }
        return result.toArray(new Integer[0]);
    }

    /**
     * 下划线转驼峰
     */
    public static String lineToHump(String str) {
        if (!str.contains("-") && !str.contains("_")) {
            return str;
        }
        str = str.toLowerCase();
        Matcher matcher = LINE_PATTERN.matcher(str);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(sb, matcher.group(1).toUpperCase());
        }
        matcher.appendTail(sb);
        return sb.toString();
    }

    /**
     * 驼峰转下划线,效率比上面高
     */
    public static String humpToLine(String str) {
        Matcher matcher = HUMP_PATTERN.matcher(str);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            matcher.appendReplacement(sb, "_" + matcher.group(0).toLowerCase());
        }
        matcher.appendTail(sb);
        return sb.toString();
    }

    public static boolean isNotBlank(CharSequence str) {
        return !isBlank(str);
    }

    public static boolean isBlank(CharSequence str) {
        int strLen;
        if (str == null || (strLen = str.length()) == 0) {
            return true;
        }
        for (int i = 0; i < strLen; i++) {
            if ((!Character.isWhitespace(str.charAt(i)))) {
                return false;
            }
        }
        return true;
    }

    public static String toUpperCase(String input) {
        return toUpperCase(input, UPPERCASE_MAP::get, e -> false);
    }

    public static String toUpperCase(String input, Function<Character, Character> replaceMap, Predicate<Character> predicate) {
        if (input == null || input.isEmpty()) {
            return input;
        }
        //半角转全角：
        char[] chars = null;
        for (int i = 1, len = input.length() - 1; i < len; i++) {
            Character character = replaceMap.apply(input.charAt(i));
            if (character == null) {
                continue;
            }
            if (chars == null) {
                chars = input.toCharArray();
            }
            char prev = input.charAt(i - 1);
            char next = input.charAt(i + 1);
            if (isEnglishLetter(prev, predicate) || Character.isDigit(prev) ||
                    isEnglishLetter(next, predicate) || Character.isDigit(next)) {
                continue;
            }
            chars[i] = character;
        }
        if (chars != null) {
            return new String(chars);
        } else {
            return input;
        }
    }

    public static boolean isEnglishLetter(char input, Predicate<Character> predicate) {
        boolean b = (input >= 'a' && input <= 'z')
                || (input >= 'A' && input <= 'Z');
        if (b) {
            return true;
        }
        return predicate.test(input);
    }
}
