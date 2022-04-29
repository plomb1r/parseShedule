import app.Frequency;
import app.Parser;
import exceptions.ElementNotFoundException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.assertThat;
import static org.assertj.core.api.Assertions.assertThatThrownBy;

class ParserTest {

    @Test
    public void findContentTest() throws ElementNotFoundException, IOException {
        String html = readFile("find_content.html");
        var elementResult = Parser.findContent(html);
        assertThat(elementResult.text()).isEqualTo("the content you are looking for");
    }

    @Test
    public void findContentShouldThrowErrorTest() throws ElementNotFoundException {
        String html = "<html></html>";
        assertThatThrownBy(() -> Parser.findContent(html))
                .isInstanceOf(ElementNotFoundException.class)
                .hasMessage("Element not found with " + Parser.TABLE_CSS_QUERY);
    }

    @Test
    public void selectorTableResultTest() throws IOException {
        String html = readFile("selectorTableResult.html");
        Document doc = Jsoup.parse(html);
        Element e = doc.selectFirst("table.selectors");
        var res = Parser.parseSelectorTable(e);
        assertThat(res.dataTerm).isEqualTo("Осенний");
        assertThat(res.dataYear).isEqualTo("2021 - 2022");
        assertThat(res.departmentName).isEqualTo("Прикладной математики");
    }
    @Test
    public void parseTimetableSelectorTest() throws IOException {
        String html = readFile("parseTimetableSelectorWithSingleTable.html");
        Document doc = Jsoup.parse(html);
        Element e = doc.selectFirst("table.timetable");
        var timetableList = Parser.parseTimetableSelector(e);
        assertThat(timetableList.size()).isEqualTo(1);
        var res = timetableList.get(0);
        assertThat(res).isNotNull();
        assertThat(res.day).isEqualTo("Вторник");
        assertThat(res.fio).isEqualTo("Баранович А.Е.");
        assertThat(res.frequency).isEqualTo(Frequency.WEEKLY);
        assertThat(res.time).isEqualTo("15:35 - 17:05");
        assertThat(res.group).isEqualTo("3бЛ 3");
        assertThat(res.place).isEqualTo("Дистанционно");
        assertThat(res.subject).isEqualTo("Современный аппарат логистического управления");
        assertThat(res.typeOfSubject).isEqualTo("Лекции");
    }

    public static String readFile(String fileName) throws IOException {
        return Files.readString(Path.of("src/test/resources/" + fileName));
    }
}