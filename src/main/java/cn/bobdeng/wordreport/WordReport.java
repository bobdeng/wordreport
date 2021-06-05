package cn.bobdeng.wordreport;

import lombok.extern.java.Log;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.impl.values.XmlValueDisconnectedException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.logging.Level;

@Log
class WordReport {
    private final ReportData reportData;

    WordReport(ReportData reportData) {
        this.reportData = reportData;
    }

    byte[] output(byte[] templateFile) throws IOException {
        XWPFDocument srcDoc = new XWPFDocument(new ByteArrayInputStream(templateFile));
        srcDoc.getHeaderList().forEach(header -> runBody(header.getTables(), header.getParagraphs()));
        srcDoc.getFooterList().forEach(footer -> runBody(footer.getTables(), footer.getParagraphs()));
        runBody(srcDoc.getTables(), srcDoc.getParagraphs());
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        srcDoc.write(output);
        return output.toByteArray();
    }

    private void runBody(List<XWPFTable> tables, List<XWPFParagraph> paragraphs) {
        tables.forEach(this::doWithTable);
        runParagraphs(paragraphs);
    }

    private void runParagraphs(List<XWPFParagraph> paragraphs) {
        this.run(paragraphs, paragraph -> {
            XWPFRun firstRun = paragraph.getRuns().get(0);
            ParagraphParse paragraphParse = new ParagraphParse(paragraph);
            paragraphParse.forEach(fragment -> fragment.create(paragraph, firstRun, reportData));
        });
    }

    private void run(List<XWPFParagraph> paragraphs, Consumer<XWPFParagraph> paragraphConsumer) {
        List<XWPFParagraph> newParagraphs = new ArrayList<>(paragraphs);
        newParagraphs
                .stream()
                .filter(paragraph -> !paragraph.getRuns().isEmpty())
                .filter(paragraph -> new ParagraphParse(paragraph).hasPlaceholder())
                .forEach(paragraph -> {
                    int oldRuns = paragraph.getRuns().size();
                    paragraphConsumer.accept(paragraph);
                    try {
                        Utils.run(() -> paragraph.removeRun(0), oldRuns);
                    } catch (XmlValueDisconnectedException e) {
                        //说明已经被删除，不必继续删了
                    }
                });
    }

    private void replaceCellParagraph(int rowIndex, XWPFTableCell cell) {
        run(cell.getParagraphs(), paragraph -> {
            ParagraphParse paragraphParse = new ParagraphParse(paragraph);
            XWPFRun firstRun = paragraph.getRuns().get(0);
            paragraphParse.forEach(fragment -> fragment.create(paragraph, firstRun, reportData, rowIndex));
        });
    }


    private void doWithTable(XWPFTable table) {
        List<Integer> rowIndexNeedRemove = new ArrayList<>();
        for (int i = 0; i < table.getRows().size(); i++) {
            appendValueToRow(table, rowIndexNeedRemove, i);
        }
        rowIndexNeedRemove.forEach(table::removeRow);
    }

    private void appendValueToRow(XWPFTable table, List<Integer> rowIndexNeedRemove, int i) {
        XWPFTableRow row = table.getRows().get(i);
        if (isArray(row)) {
            appendArrayValueInTable(table, rowIndexNeedRemove, i, row);
            return;
        }
        replaceSingleValueInTable(row);
    }

    private void appendArrayValueInTable(XWPFTable table, List<Integer> rowIndexNeedRemove, int i, XWPFTableRow row) {
        appendArrayValuesInTable(table, row, i);
        rowIndexNeedRemove.add(i);
    }

    private boolean isArray(XWPFTableRow row) {
        return row.getTableCells()
                .stream()
                .anyMatch(this::isArray);
    }

    private boolean isArray(XWPFTableCell xwpfTableCell) {
        return xwpfTableCell.getParagraphs()
                .stream()
                .map(ParagraphParse::new)
                .anyMatch(ParagraphParse::hasArray);
    }

    private void replaceSingleValueInTable(XWPFTableRow row) {
        row.getTableCells()
                .stream()
                .map(XWPFTableCell::getParagraphs)
                .forEach(this::runParagraphs);
    }

    private void appendArrayValuesInTable(XWPFTable table, XWPFTableRow row, int rowIndex) {
        int indexOfList = 0;
        while (true) {
            if (!appendRow(table, row, rowIndex, indexOfList)) {
                return;
            }
            indexOfList++;
        }

    }

    private boolean appendRow(XWPFTable table, XWPFTableRow row, int rowIndex, int indexOfList) {
        try {
            XWPFTableRow newRow = cloneRow(table, row);
            boolean hasValue = replaceCellsInRow(newRow, indexOfList);
            if (!hasValue) {
                return false;
            }
            table.addRow(newRow, rowIndex + indexOfList + 1);
        } catch (Exception e) {
            log.log(Level.WARNING, "", e);
        }
        return true;
    }

    private XWPFTableRow cloneRow(XWPFTable table, XWPFTableRow row) throws XmlException, IOException {
        CTRow ctrow = CTRow.Factory.parse(row.getCtRow().newInputStream());
        return new XWPFTableRow(ctrow, table);
    }

    private boolean replaceCellsInRow(XWPFTableRow row, int rowIndex) {
        if (!hadMoreValue(row, rowIndex)) {
            return false;

        }
        row.getTableCells()
                .forEach(cell -> replaceCellParagraph(rowIndex, cell));
        return true;
    }


    private boolean hadMoreValue(XWPFTableRow row, int rowIndex) {
        return row.getTableCells()
                .stream()
                .anyMatch(cell1 -> hadMoreValue(rowIndex, cell1));
    }

    private boolean hadMoreValue(int rowIndex, XWPFTableCell cell) {
        return cell.getParagraphs()
                .stream()
                .map(ParagraphParse::new)
                .anyMatch(paragraphParse -> paragraphParse.hasValue(reportData, rowIndex));
    }

}
