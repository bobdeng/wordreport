package cn.bobdeng.wordreport;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;


public interface TemplateContent {

    TemplateContent combine(TemplateContent content);

    void append(XWPFParagraph paragraph, XWPFRun firstRun, Fragment fragment);

    boolean append(XWPFParagraph paragraph, XWPFRun firstRun, int rowIndex);

    boolean hasValue(int index);
}
