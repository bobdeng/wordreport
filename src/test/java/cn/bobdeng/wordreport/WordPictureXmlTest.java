package cn.bobdeng.wordreport;

import cn.bobdeng.testtools.TestResource;
import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

class WordPictureXmlTest {
    @Test
    void should_return_right_xml() throws IOException {
        WordPictureXml wordPictureXml=new WordPictureXml(1,"2",100,200);
        System.out.println(wordPictureXml.toXml());
        assertThat(wordPictureXml.toXml(), is(new TestResource(this,"picture.xml").readString()));
    }
}
