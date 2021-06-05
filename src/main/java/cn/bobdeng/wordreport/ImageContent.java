package cn.bobdeng.wordreport;

import lombok.Getter;
import lombok.extern.java.Log;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.Serializable;
import java.util.Collections;
import java.util.List;
import java.util.UUID;
import java.util.logging.Level;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Getter
@Log
public class ImageContent implements TemplateContent, Serializable {
    private final List<byte[]> images;
    //no use, just for json converter
    private boolean image = true;

    public ImageContent(List<byte[]> images) {
        this.images = images;
    }

    public ImageContent(byte[] image) {
        this(Collections.singletonList(image));
    }

    private double fontSizeToPx(double fontSize) {
        return fontSize * 4 / 3;
    }

    @Override
    public TemplateContent combine(TemplateContent content) {
        ImageContent imageContent = (ImageContent) content;
        return new ImageContent(Stream.concat(images.stream(), imageContent.images.stream()).collect(Collectors.toList()));
    }

    @Override
    public void append(XWPFParagraph paragraph, XWPFRun firstRun, Fragment fragment) {
        images.forEach(imageData -> appendImageData(paragraph, firstRun, imageData));
    }

    private void appendImageData(XWPFParagraph paragraph, XWPFRun firstRun, byte[] imageData) {
        try {
            XWPFRun run = paragraph.createRun();
            Utils.copyRunStyle(run, firstRun);
            insertImage(imageData, run);
        } catch (Exception e) {
            log.log(Level.WARNING, "", e);
        }
    }

    private void insertImage(byte[] imageData, XWPFRun run) throws IOException, InvalidFormatException {
        double imageHeight = calculateImageHeight(run);
        ByteArrayInputStream pictureData = new ByteArrayInputStream(imageData);
        BufferedImage bufferedImage = ImageIO.read(pictureData);
        double imageWith = bufferedImage.getWidth() * imageHeight / bufferedImage.getHeight();
        run.addPicture(new ByteArrayInputStream(imageData), Document.PICTURE_TYPE_JPEG, UUID.randomUUID().toString(), Units.toEMU(imageWith), Units.toEMU(imageHeight));
    }

    private double calculateImageHeight(XWPFRun run) {
        double fontSize = run.getFontSize();
        if (fontSize == -1) {
            fontSize = 10.5f;
        }
        return fontSizeToPx(fontSize);
    }

    @Override
    public boolean append(XWPFParagraph paragraph, XWPFRun firstRun, int rowIndex) {
        if (rowIndex >= images.size()) {
            return false;
        }
        this.appendImageData(paragraph, firstRun, images.get(rowIndex));
        return true;
    }

    @Override
    public boolean hasValue(int index) {
        return images.size() > index;
    }

    public int size() {
        return images.size();
    }
}
