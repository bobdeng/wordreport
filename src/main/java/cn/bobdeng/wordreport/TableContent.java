package cn.bobdeng.wordreport;

import lombok.Getter;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.List;

import static cn.bobdeng.wordreport.Utils.averageTableCellWidth;

@Getter
public class TableContent implements TemplateContent {
    private final List<String> titles;
    private final List<List<String>> lines;

    public TableContent(List<String> titles, List<List<String>> lines) {
        this.titles = titles;
        this.lines = lines;
    }

    @Override
    public TemplateContent combine(TemplateContent content) {
        return null;
    }

    @Override
    public void append(XWPFParagraph paragraph, XWPFRun firstRun, Fragment fragment) {
        BodyType partType = paragraph.getPartType();
        if (partType == BodyType.TABLECELL) {
            appendInTable(paragraph);
            return;
        }
        createTableAndAppend(paragraph);
    }

    private void createTableAndAppend(XWPFParagraph paragraph) {
        XWPFTable table = paragraph.getDocument().insertNewTbl(paragraph.getCTP().newCursor());
        buildTable(paragraph, table);
    }

    private void buildTable(XWPFParagraph paragraph, XWPFTable table) {
        expandTable(table, titles.size() - 1);
        setValuesOfTable(table);
        CTTcPr tcPrOfFirstCell = table.getRow(0).getCell(0).getCTTc().getTcPr();
        averageTableCellWidth(paragraph, titles, table);
        setAllBordersToFirstCell(table, tcPrOfFirstCell);
        table.removeRow(0);
    }

    private void setAllBordersToFirstCell(XWPFTable table, CTTcPr tcPrOfFirstCell) {
        table.getRows().forEach(xwpfTableRow -> xwpfTableRow.getTableCells().forEach(cell -> {
            CTTc ctTc = cell.getCTTc();
            CTTcPr ctTcTcPr = ctTc.getTcPr();
            if (ctTcTcPr != null && tcPrOfFirstCell != null) {
                ctTcTcPr.setTcBorders(tcPrOfFirstCell.getTcBorders());
            }
        }));
    }

    private void appendInTable(XWPFParagraph paragraph) {
        XWPFTableCell part = (XWPFTableCell) paragraph.getBody();
        XWPFTable table = part.getTableRow().getTable();
        buildTable(paragraph, table);
    }

    private void expandTable(XWPFTable table, int rows) {
        for (int i = 0; i < rows; i++) {
            table.addNewCol();
        }
        for (int i = 0; i <= lines.size(); i++) {
            table.createRow();
        }
    }

    private void setValuesOfTable(XWPFTable table) {
        int from = 1;
        XWPFTableRow rowTitle = table.getRow(from);
        setTableLine(rowTitle, titles);
        for (int i = 0; i < lines.size(); i++) {
            XWPFTableRow rowLine = table.getRow(i + 1 + from);
            setTableLine(rowLine, lines.get(i));
        }
    }

    private void setTableLine(XWPFTableRow rowTitle, List<String> titles) {
        for (int i = 0; i < titles.size(); i++) {
            rowTitle.getCell(i).setText(titles.get(i));
        }
    }

    @Override
    public boolean append(XWPFParagraph paragraph, XWPFRun firstRun, int rowIndex) {
        return false;
    }

    @Override
    public boolean hasValue(int index) {
        return true;
    }
}
