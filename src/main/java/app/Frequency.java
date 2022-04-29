package app;

public enum Frequency {
    NUMERATOR("Числитель"),
    DENOMINATOR("Знаменатель"),
    WEEKLY("Еженедельно");
    public final String htmlValue;

    Frequency(String htmlValue) {
        this.htmlValue = htmlValue;
    }
    public static Frequency byHtmlValue(String htmlValue){
        for (Frequency frequency : Frequency.values()) {
            if (frequency.htmlValue.equals(htmlValue)){
                return frequency;
            }
        }
        throw new RuntimeException("Not found type of week");
    }
}
