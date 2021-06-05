package cn.bobdeng.wordreport;

import lombok.Getter;
import org.apache.poi.xwpf.usermodel.*;

import java.util.Arrays;
import java.util.List;

import static cn.bobdeng.wordreport.Placeholder.PLACEHOLDER_BEGIN;
import static cn.bobdeng.wordreport.Placeholder.PLACEHOLDER_END;
import static cn.bobdeng.wordreport.Utils.averageTableCellWidth;

@Getter
public class PlaceholderFragment implements Fragment {
    private final String fragment;

    PlaceholderFragment(String fragment) {

        this.fragment = fragment;
    }

    @Override
    public void create(XWPFParagraph paragraph, XWPFRun firstRun, ReportData reportData) {
        if (fragment.startsWith("≮table:")) {
            insertFullTable(paragraph, reportData);
            return;
        }
        TemplateContent content = reportData.getContent(getPlaceholder());
        if (content == null) {
            return;
        }
        content.append(paragraph, firstRun, this);
    }

    private String getPlaceholder() {
        if (fragment.contains("{")) {
            return fragment.substring(0, fragment.indexOf('{')) + PLACEHOLDER_END;
        }
        return fragment;
    }

    private void insertFullTable(XWPFParagraph paragraph, ReportData reportData) {
        List<String> fields = getTableFields();
        XWPFTable table = paragraph.getDocument().createTable(1, fields.size());
        averageTableCellWidth(paragraph, fields, table);
        setTableHeader(fields, table);
        setTableContent(fields, table, reportData);
    }

    private void setTableHeader(List<String> fields, XWPFTable table) {
        XWPFTableRow header = table.getRow(0);
        for (int i = 0; i < header.getTableCells().size(); i++) {
            String title = fields.get(i);
            header.getCell(i).setText(title);
        }
    }

    private void setTableContent(List<String> fields, XWPFTable table, ReportData reportData) {
        int col = 0;
        while (true) {
            XWPFTableRow row = table.createRow();
            if (!insertRow(fields, col, row, reportData)) {
                break;
            }
            col++;
        }
        table.removeRow(table.getRows().size() - 1);
    }

    private boolean insertRow(List<String> fields, int col, XWPFTableRow row, ReportData reportData) {
        boolean hasMore = false;
        for (int i = 0; i < row.getTableCells().size(); i++) {
            if (setCellInRow(fields, col, row, i, reportData)) {
                hasMore = true;
            }
        }
        return hasMore;
    }

    @Override
    public boolean isDistinct() {
        return fragment.contains("distinct");
    }

    private boolean setCellInRow(List<String> fields, int col, XWPFTableRow row, int cellIndex, ReportData reportData) {
        String fieldName = fields.get(cellIndex);
        TemplateContent content = reportData.getContent(PLACEHOLDER_BEGIN + fieldName + PLACEHOLDER_END);
        XWPFTableCell cell = row.getCell(cellIndex);
        if (content == null) {
            return false;
        }
        XWPFParagraph paragraph = cell.getParagraphArray(0);
        return content.append(paragraph, paragraph.createRun(), col);
    }

    private List<String> getTableFields() {
        return Arrays.asList(this.fragment.replace(PLACEHOLDER_BEGIN + "table:", "")
                .replace(PLACEHOLDER_END, "")
                .split(","));
    }

    @Override
    public boolean create(XWPFParagraph paragraph, XWPFRun firstRun, ReportData reportData, int rowIndex) {
        TemplateContent content = reportData.getContent(fragment);
        if (content == null) {
            return false;
        }
        return content.append(paragraph, firstRun, rowIndex);
    }

    @Override
    public boolean hasNewLine() {
        return fragment.contains("newLine");
    }
}
