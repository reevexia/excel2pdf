package com.reeve.dict;

/**
 * @Author Reeve
 * @Date 2021/6/25 16:51
 */
public enum LetterValue {

    A("A", 0),
    B("B", 1),
    C("C", 2),
    D("D", 3),
    E("E", 4),
    F("F", 5),
    G("G", 6),
    H("H", 7),
    I("I", 8),
    J("J", 9),
    K("K", 10),
    L("L", 11),
    M("M", 12),
    N("N", 13),
    O("O", 14),
    P("P", 15),
    Q("Q", 16),
    R("R", 17),
    S("S", 18),
    T("T", 19),
    U("U", 20),
    V("V", 21),
    W("W", 22),
    X("X", 23),
    Y("Y", 24),
    Z("Z", 25);

    private String code;
    private int value;

    LetterValue(String code, int value) {
        this.code = code;
        this.value = value;
    }

    public String getCode() {
        return code;
    }

    public int getValue() {
        return value;
    }

    public static LetterValue obj2Enum(Object obj) {
        if (null == obj) {
            return null;
        }
        for (LetterValue enumInstance : LetterValue.values()) {
            if (enumInstance.getCode().equals(obj)) {
                return enumInstance;
            }
        }
        for (LetterValue enumInstance : LetterValue.values()) {
            if (enumInstance.getValue() == (Integer) obj) {
                return enumInstance;
            }
        }
        return null;
    }

    public static boolean isValid(String code) {
        return null != obj2Enum(code);
    }

    public boolean equalsIgnoreCase(String target) {
        return this.getCode().equalsIgnoreCase(target);
    }
}
