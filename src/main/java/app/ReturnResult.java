package app;

import org.jsoup.nodes.Element;

public class ReturnResult {
    public final Element selectorsElement;
    public final Element timetableElement;
    public ReturnResult(Element selectorsElement, Element timetableElement) {
        this.selectorsElement = selectorsElement;
        this.timetableElement = timetableElement;
    }
}