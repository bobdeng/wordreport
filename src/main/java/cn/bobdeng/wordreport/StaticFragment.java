package cn.bobdeng.wordreport;

import lombok.Getter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

@Getter
public class StaticFragment implements Fragment {
    private final String fragment;

    StaticFragment(String fragment) {

        this.fragment = fragment;
    }

    @Override
    public void create(XWPFParagraph paragraph, XWPFRun firstRun, ReportData reportData) {
        XWPFRun run = paragraph.createRun();
        Utils.copyRunStyle(run, firstRun);
        run.setText(fragment);
    }

    @Override
    public boolean create(XWPFParagraph paragraph, XWPFRun firstRun, ReportData reportData, int rowIndex) {
        XWPFRun run = paragraph.createRun();
        Utils.copyRunStyle(run, firstRun);
        run.setText(fragment);
        return false;
    }

    @Override
    public boolean hasNewLine() {
        return false;
    }
}
