package cn.bobdeng.wordreport;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

public class WordCombine {
    private final List<byte[]> wordFiles;

    public WordCombine(List<byte[]> wordFiles) {
        this.wordFiles = wordFiles;
    }

    public byte[] combine() throws IOException, InvalidFormatException {
        XWPFDocument result = new XWPFDocument(OPCPackage.open(new ByteArrayInputStream(wordFiles.get(0))));
        appendOtherDocument(result);
        return documentOutput(result);
    }

    private byte[] documentOutput(XWPFDocument result) throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        result.write(output);
        return output.toByteArray();
    }

    private void appendOtherDocument(XWPFDocument result) {
        wordFiles.stream().skip(1).forEach(wordFile -> {
            try {
                XWPFDocument append = new XWPFDocument(OPCPackage.open(new ByteArrayInputStream(wordFile)));
                appendBody(result, append);
            } catch (Exception e) {
                throw new RuntimeException("文件格式错误");
            }
        });
    }

    public static void appendBody(XWPFDocument src, XWPFDocument append) {
        src.createParagraph().setPageBreak(true);
        CTBody newBody = src.getDocument().addNewBody();
        newBody.addNewSectPr();
        newBody.set(append.getDocument().getBody());

    }
}
