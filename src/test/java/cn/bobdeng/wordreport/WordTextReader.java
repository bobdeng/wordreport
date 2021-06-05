package cn.bobdeng.wordreport;

import lombok.Getter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.function.Consumer;

public class WordTextReader {
    @Getter
    private final String text;

    public WordTextReader(File file) throws IOException {
        StringBuffer stringBuffer = new StringBuffer();
        Consumer<XWPFTable> xwpfTableConsumer = xwpfTable -> stringBuffer.append(readTable(xwpfTable));
        Consumer<XWPFParagraph> xwpfParagraphConsumer = paragraph -> {
            stringBuffer.append(readParagraph(paragraph));
        };
        readDoc(file, xwpfTableConsumer, xwpfParagraphConsumer);
        this.text = stringBuffer.toString();
    }

    private void readDoc(File file, Consumer<XWPFTable> xwpfTableConsumer, Consumer<XWPFParagraph> xwpfParagraphConsumer) throws IOException {
        XWPFDocument srcDoc = new XWPFDocument(new FileInputStream(file));
        srcDoc.getHeaderList().forEach(xwpfHeader -> {
            xwpfHeader.getTables().forEach(xwpfTableConsumer);
            xwpfHeader.getParagraphs().forEach(xwpfParagraphConsumer);
        });
        srcDoc.getFooterList().forEach(xwpfHeader -> {
            xwpfHeader.getTables().forEach(xwpfTableConsumer);
            xwpfHeader.getParagraphs().forEach(xwpfParagraphConsumer);
        });
        srcDoc.getTables().forEach(xwpfTableConsumer);
        srcDoc.getParagraphs().forEach(xwpfParagraphConsumer);

    }

    private String readTable(XWPFTable xwpfTable) {
        StringBuffer stringBuffer = new StringBuffer();
        xwpfTable.getRows().forEach(xwpfTableRow -> {
            xwpfTableRow.getTableCells().forEach(xwpfTableCell -> {
                xwpfTableCell.getParagraphs().forEach(paragraph -> stringBuffer.append(readParagraph(paragraph)));
            });
        });
        return stringBuffer.toString();
    }

    private String readParagraph(XWPFParagraph paragraph) {
        StringBuffer stringBuffer = new StringBuffer();
        paragraph.getRuns().forEach(xwpfRun -> {
            stringBuffer.append(xwpfRun.toString());
            if (!xwpfRun.getEmbeddedPictures().isEmpty()) {
                readPicture(stringBuffer, xwpfRun);
            }
        });
        return stringBuffer.toString();
    }

    private void readPicture(StringBuffer stringBuffer, org.apache.poi.xwpf.usermodel.XWPFRun xwpfRun) {
        XWPFPicture xwpfPicture = xwpfRun.getEmbeddedPictures().get(0);
        byte[] data = xwpfPicture.getPictureData().getData();
        String base64 = Base64.getEncoder().encodeToString(data);
        stringBuffer.append("image:" + base64);
    }
}
