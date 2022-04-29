package app;

import exceptions.ElementNotFoundException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.util.ArrayList;
import java.util.List;


public class Parser {
    public final static String TABLE_CSS_QUERY = "div#result";
    public final static String SELECTORS_ELEMENT = "table.selectors";
    public final static String TIMETABLE_ELEMENT = "table.timetable";

    public static Element findContent(String html) throws ElementNotFoundException {
        Document doc = Jsoup.parse(html);
        Element element = doc.selectFirst(TABLE_CSS_QUERY);
        if (element == null) {
            throw new ElementNotFoundException("Element not found with " + TABLE_CSS_QUERY);
        }
        return element;
    }

    public static ReturnResult getTables(Element tableContainer) {
        var selectorElement = tableContainer.selectFirst("table.selectors");
        if (selectorElement == null) {
            throw new ElementNotFoundException("Element not found with " + SELECTORS_ELEMENT);
        }
        var timetableElement = tableContainer.selectFirst("table.timetable");
        if (timetableElement == null) {
            throw new ElementNotFoundException("Element not found with " + TIMETABLE_ELEMENT);
        }
        return new ReturnResult(selectorElement, timetableElement);
    }

    public static SelectorTableResult parseSelectorTable(Element selectorElement) {
        var thIter = selectorElement.select("th").iterator();
        var tdIter = selectorElement.select("td").iterator();
        String dY = null;
        String dT = null;
        String dN = null;
        while (thIter.hasNext() && tdIter.hasNext()) {
            String thKey = thIter.next().text();
            String tdValue = tdIter.next().text();
            switch (thKey) {
                case "Учебный год:":
                    dY = tdValue;
                    break;
                case "Семестр:":
                    dT = tdValue;
                    break;
                case "Кафедра:":
                    dN = tdValue;
                    break;
                case "Форма обучения:":
                    break;
                default:
                    throw new IllegalArgumentException("Invalid htmlTable");
            }
        }
        if (thIter.hasNext() || tdIter.hasNext()) {
            throw new RuntimeException("Different number of elements");
        }
        return new SelectorTableResult(dY, dT, dN);
    }

    public static List<TimeTableResult> parseTimetableSelector(Element timetableElement) {
        var trIter = timetableElement.select("tr").iterator();
        List<TimeTableResult> resultList = new ArrayList<>();
        String fio = null;
        String day = null;
        while (trIter.hasNext()) {
            var trKey = trIter.next();
            var trChildren = trKey.children();
            var firstTrChildren = trChildren.get(0);
            var tag = firstTrChildren.tag().getName();
            if (trChildren.size() == 1) {
                if (tag.equals("th")) {
                    day = trKey.text();
                } else if (tag.equals("td")) {
                    fio = trKey.text();
                } else {
                    throw new RuntimeException("Unknown tag in timetable");
                }
            } else if (trChildren.size() == 6) {
                if (tag.equals("td") && firstTrChildren.select("b").size() == 0) {
                    String time = trChildren.get(0).text();
                    Frequency frequency = Frequency.byHtmlValue(trChildren.get(1).text());
                    String place = trChildren.get(2).text();
                    String group = trChildren.get(3).text();
                    String subject = trChildren.get(4).text();
                    String typeOfSubject = trChildren.get(5).text();
                    TimeTableResult tr = new TimeTableResult(fio, day, time, frequency, place, group, subject, typeOfSubject);
                    resultList.add(tr);
                }
            } else {
                throw new RuntimeException("Unexpected number of <tr>");
            }
        }
        return resultList;
    }
}




