package cn.bobdeng.wordreport;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;
import java.util.List;
import java.util.Optional;

class Utils {
    private Utils() {
    }

    static void copyRunStyle(XWPFRun run, XWPFRun firstRun) {
        copyFontSize(run, firstRun);
        copyFontColor(run, firstRun);
        copyFontStyle(run, firstRun);
    }

    private static void copyFontStyle(XWPFRun run, XWPFRun firstRun) {
        run.setFontFamily(getFontFamily(firstRun));
        run.setBold(firstRun.isBold());
        run.setItalic(firstRun.isItalic());
        run.setUnderline(firstRun.getUnderline());
    }

    private static void copyFontColor(XWPFRun run, XWPFRun firstRun) {
        if (firstRun.getColor() != null) {
            run.setColor(firstRun.getColor());
        }
    }

    private static void copyFontSize(XWPFRun run, XWPFRun firstRun) {
        if (firstRun.getFontSize() > 0) {
            run.setFontSize(firstRun.getFontSize());
        }
    }

    private static String getFontFamily(XWPFRun run) {
        return Optional.ofNullable(run.getFontFamily())
                .orElse("宋体");
    }

    static void run(Runnable runnable, int times) {
        for (int i = 0; i < times; i++) {
            runnable.run();
        }
    }

    static int getDocumentWidth(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) return 8600;
        CTPageSz pageSize = sectPr.getPgSz();
        if (pageSize == null) return 8600;
        return getPageWidth(sectPr, pageSize);
    }

    private static int getPageWidth(CTSectPr sectPr, CTPageSz pageSize) {
        double width = Math.round(pageSize.getW().doubleValue());
        CTPageMar pageMargin = sectPr.getPgMar();
        double pageMarginLeft = pageMargin.getLeft().doubleValue();
        double pageMarginRight = pageMargin.getRight().doubleValue();
        double effectivePageWidth = width - pageMarginLeft - pageMarginRight;
        return (int) effectivePageWidth;
    }

    public static void averageTableCellWidth(XWPFParagraph paragraph, List<String> fields, XWPFTable table) {
        int width = getDocumentWidth(paragraph.getDocument()) / fields.size();
        table.getRows().forEach(row -> setRowWidth(width, row));
    }

    private static void setRowWidth(int width, XWPFTableRow row) {
        row.getTableCells().forEach(cell -> setCellWidth(width, cell));
    }

    private static void setCellWidth(int width, XWPFTableCell cell) {
        CTTblWidth cellWidth = cell.getCTTc().addNewTcPr().addNewTcW();
        CTTcPr pr = cell.getCTTc().addNewTcPr();
        pr.addNewNoWrap();
        cellWidth.setW(BigInteger.valueOf(width));
    }

}
