package cn.bobdeng.wordreport;

import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

public class ParagraphParseTest {
    @Test
    public void 切分不同领域() {
        ParagraphParse paragraphParse = new ParagraphParse("123≮考核序号≯、≮要素类型≯ 12345");
        List<Fragment> fragments = new ArrayList<>();
        paragraphParse.forEach(fragment -> {
            fragments.add(fragment);
        });
        assertThat(fragments.size(), is(5));
    }

}
