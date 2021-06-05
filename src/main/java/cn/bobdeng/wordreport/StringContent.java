package cn.bobdeng.wordreport;

import lombok.EqualsAndHashCode;
import lombok.Getter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Getter
@EqualsAndHashCode
public class StringContent implements TemplateContent {
    private final List<String> values;

    public StringContent(List<String> values) {
        this.values = values;
    }

    public StringContent(String value) {
        this(Collections.singletonList(value));
    }

    private String getValue(int index) {
        if (exceedMax(index)) {
            return "";
        }
        return values.get(index);
    }


    @Override
    public TemplateContent combine(TemplateContent content) {
        StringContent stringArrayContent = (StringContent) content;
        List<String> newValues = Stream.concat(getValues().stream(), stringArrayContent.getValues().stream())
                .collect(Collectors.toList());
        return new StringContent(newValues);
    }

    @Override
    public void append(XWPFParagraph paragraph, XWPFRun firstRun, Fragment fragment) {
        if (fragment.isDistinct()) {
            List<String> distinctValues = this.values.stream()
                    .distinct().collect(Collectors.toList());
            appendValues(paragraph, firstRun, fragment, distinctValues);
            return;
        }
        appendValues(paragraph, firstRun, fragment, this.values);
    }

    private void appendValues(XWPFParagraph paragraph, XWPFRun firstRun, Fragment fragment, List<String> stringList) {
        if (fragment.hasNewLine()) {
            appendStringsSplitWithNewLine(paragraph, firstRun, stringList);
            return;
        }
        appendStringsWithJoined(paragraph, firstRun, stringList);
    }

    private void appendStringsWithJoined(XWPFParagraph paragraph, XWPFRun firstRun, List<String> stringList) {
        createRun(paragraph, firstRun, stringList.stream()
                .filter(Objects::nonNull)
                .collect(Collectors.joining(" ")));
    }

    private void appendStringsSplitWithNewLine(XWPFParagraph paragraph, XWPFRun firstRun, List<String> stringList) {
        stringList.stream().filter(Objects::nonNull)
                .forEach(value -> {
                    XWPFRun run = createRun(paragraph, firstRun, value);
                    run.addBreak();
                });
    }

    private XWPFRun createRun(XWPFParagraph paragraph, XWPFRun firstRun, String value) {
        XWPFRun run = paragraph.createRun();
        Utils.copyRunStyle(run, firstRun);
        run.setText(value);
        return run;
    }

    @Override
    public boolean append(XWPFParagraph paragraph, XWPFRun firstRun, int rowIndex) {
        createRun(paragraph, firstRun, getValue(rowIndex));
        return !exceedMax(rowIndex);
    }

    private boolean exceedMax(int index) {
        return index >= values.size();
    }

    public boolean hasValue(int index) {
        return !exceedMax(index);
    }
}
