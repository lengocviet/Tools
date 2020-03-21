package Regrex;

public class TestRegrex <T> {

    public static <T> String $(String A, T B)
    {
        return A + B.toString();
    }
}
