package cn.bobdeng.wordreport;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.function.Consumer;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class ParagraphParse {

    private final String[] fragments;

    public ParagraphParse(String content) {
        content = content.replaceAll(Placeholder.PLACEHOLDER_BEGIN, "\n" + Placeholder.PLACEHOLDER_BEGIN);
        content = content.replaceAll(Placeholder.PLACEHOLDER_END, Placeholder.PLACEHOLDER_END + "\n");
        Pattern pattern = Pattern.compile("\\n");
        this.fragments = pattern.split(content);
    }

    public ParagraphParse(XWPFParagraph paragraph) {
        this(paragraph.getRuns().stream()
                .map(XWPFRun::toString)
                .collect(Collectors.joining()));
    }

    public void forEach(Consumer<Fragment> fragmentConsumer) {
        Stream.of(fragments).forEach(fragment -> {
            if (isPlaceholder(fragment)) {
                fragmentConsumer.accept(new PlaceholderFragment(fragment));
                return;
            }
            fragmentConsumer.accept(new StaticFragment(fragment));
        });
    }

    private boolean isPlaceholder(String fragment) {
        return fragment.contains(Placeholder.PLACEHOLDER_BEGIN);
    }

    public boolean hasArray() {
        return Stream.of(fragments)
                .filter(this::isPlaceholder)
                .anyMatch(fragment -> fragment.contains("[]"));
    }

    public boolean hasValue(ReportData reportData, int index) {
        return Stream.of(fragments)
                .filter(this::isPlaceholder)
                .anyMatch(fragment -> {
                    TemplateContent content = reportData.getContent(fragment);
                    if (content == null) {
                        return false;
                    }
                    return content.hasValue(index);
                });
    }

    public boolean hasPlaceholder() {
        return Stream.of(fragments)
                .anyMatch(this::isPlaceholder);
    }
}
