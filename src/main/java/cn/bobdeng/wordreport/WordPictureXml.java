package cn.bobdeng.wordreport;

import com.google.common.io.ByteStreams;

import java.io.IOException;
import java.io.InputStream;

public class WordPictureXml {
    private final int id;
    private final String blipId;
    private final int cx;
    private final int cy;

    public WordPictureXml(int id, String blipId, long cx, long cy) {
        this.id = id;
        this.blipId = blipId;
        this.cx = (int) cx;
        this.cy = (int) cy;
    }

    public String toXml() throws IOException {
        return readTemplate()
                .replace("{id}", String.valueOf(id))
                .replace("{blipId}", blipId)
                .replace("{cx}", String.valueOf(cx))
                .replace("{cy}", String.valueOf(cy));
    }

    private String readTemplate() throws IOException {
        InputStream xmlStream = getClass().getClassLoader().getResourceAsStream("picture.xml");
        assert xmlStream != null;
        return new String(ByteStreams.toByteArray(xmlStream));
    }
}
