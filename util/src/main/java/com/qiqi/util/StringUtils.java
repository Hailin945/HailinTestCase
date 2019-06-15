package com.qiqi.util;

import java.util.HashMap;
import java.util.Map;

/**
 * 字符串操作工具类
 *
 * @author Hailin
 */
public class StringUtils {

    /**
     * 获取Java中所有字符以及对应编码的值
     *
     * @return Java中所有字符及其对应编码值的key-value集合
     */
    public static Map<Integer, Character> getAllJavaCharacter() {
        final Map<Integer, Character> map = new HashMap<>();
        for (int i = Character.MIN_VALUE; i <= Character.MAX_VALUE; ++i) {
            map.put(i, (char) i);
            // System.out.println(i + "    " + (char) i);
        }
        return map;
    }

    /**
     * 全角字符串转换半角字符串
     *
     * <p>
     * 全半角字符之间的关系：
     * 1、半角字符是从33开始到126结束
     * 2、与半角字符对应的全角字符是从65281开始到65374结束
     * 3、其中半角的空格是32.对应的全角空格是12288
     * 4、半角和全角的关系,除空格外的字符偏移量是65248(65281-33 = 65248)
     * </p>
     *
     * @param fullWidthStr 非空的全角字符串
     * @return 半角字符串
     */
    public static String fullWidth2halfWidth(final String fullWidthStr) {
        if (null == fullWidthStr || fullWidthStr.length() <= 0) {
            return "";
        }
        char[] charArray = fullWidthStr.toCharArray();
        // 对全角字符转换的char数组遍历
        for (int i = 0; i < charArray.length; ++i) {
            int charIntValue = (int) charArray[i];
            // 如果符合转换关系,将对应下标之间减掉偏移量65248;如果是空格的话,直接做转换
            if (charIntValue >= 65281 && charIntValue <= 65374) {
                charArray[i] = (char) (charIntValue - 65248);
            } else if (charIntValue == 12288) {
                charArray[i] = (char) 32;
            }
        }
        return new String(charArray);
    }
}