package cn.bobdeng.wordreport;

import lombok.Getter;
import lombok.extern.java.Log;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import static cn.bobdeng.wordreport.Placeholder.PLACEHOLDER_BEGIN;
import static cn.bobdeng.wordreport.Placeholder.PLACEHOLDER_END;

@Log
public class ReportData {
    @Getter
    private final Map<String, TemplateContent> values = new HashMap<>();

    public void setKey(String key, TemplateContent value) {
        String keyWithPrefix = getKeyWithPrefix(key);
        values.put(keyWithPrefix, value);
    }

    private String getKeyWithPrefix(String key) {
        return PLACEHOLDER_BEGIN + key + PLACEHOLDER_END;
    }

    public void addKey(String key, TemplateContent value) {
        combineValue(getKeyWithPrefix(key), value);
    }

    public void addKey(String key, String value) {
        this.addKey(key, new StringContent(value));
    }

    public void setKey(String key, String value) {
        this.setKey(key, new StringContent(value));
    }

    public void setKey(String key, List<String> values) {
        this.setKey(key, new StringContent(values));
    }


    public byte[] output(byte[] templateFile) throws IOException {
        return new WordReport(this).output(templateFile);
    }

    TemplateContent getContent(String placeholder) {
        return Optional.ofNullable(this.values.get(placeholder))
                .orElseGet(() -> {
                    String holderNameToArray = placeholder.replaceAll(PLACEHOLDER_END, "[]" + PLACEHOLDER_END);
                    return this.values.get(holderNameToArray);
                });
    }

    public ReportData combine(ReportData reportDataB) {
        reportDataB.getValues().forEach(this::combineValue);
        return this;
    }

    private void combineValue(String key, TemplateContent value) {
        values.computeIfPresent(key, (k, presentValue) -> presentValue.combine(value));
        values.computeIfAbsent(key, k -> value);
    }

    Object getValue(String name) {
        return getContent(PLACEHOLDER_BEGIN + name + PLACEHOLDER_END);
    }

}
