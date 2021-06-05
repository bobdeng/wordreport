package cn.bobdeng.wordreport;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public interface Fragment {
    void create(XWPFParagraph paragraph, XWPFRun firstRun, ReportData reportData);

    boolean create(XWPFParagraph paragraph, XWPFRun firstRun, ReportData reportData, int rowIndex);

    boolean hasNewLine();

    default boolean isDistinct() {
        return false;
    }
}
