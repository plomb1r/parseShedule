package app;

public class TimeTableResult {
    public final String fio;
    public final String day;
    public final String time;
    public final Frequency frequency;
    public final String place;
    public final String group;
    public final String subject;
    public final String typeOfSubject;

    public TimeTableResult(String fio, String day, String time, Frequency frequency, String place, String group, String subject, String typeOfSubject) {
        this.fio = fio;
        this.day = day;
        this.time = time;
        this.frequency = frequency;
        this.place = place;
        this.group = group;
        this.subject = subject;
        this.typeOfSubject = typeOfSubject;
    }
}