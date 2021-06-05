package cn.bobdeng.wordreport;

import org.junit.Test;

import java.util.Arrays;

import static cn.bobdeng.testtools.SnapshotMatcher.snapshotMatch;
import static org.hamcrest.CoreMatchers.is;
import static org.junit.Assert.assertThat;

public class ReportDataTest {
    @Test
    public void Given有相同的Key_When合并_Then合并为数组() {
        ReportData reportDataA = new ReportData();
        reportDataA.setKey("name", new StringContent("valueA"));
        ReportData reportDataB = new ReportData();
        reportDataB.setKey("name", new StringContent("valueB"));
        reportDataA = reportDataA.combine(reportDataB);
        assertThat(reportDataA,snapshotMatch(this,"combine_same_to_array"));
    }

    @Test
    public void Given有不同的Key_When合并_Then合并两个() {
        ReportData reportDataA = new ReportData();
        reportDataA.setKey("name", new StringContent("valueA"));
        ReportData reportDataB = new ReportData();
        reportDataB.setKey("name1", new StringContent("valueB"));
        reportDataA = reportDataA.combine(reportDataB);
        assertThat(reportDataA,snapshotMatch(this,"combine_different"));
    }

    @Test
    public void Given没有值_When设置字符串_Then转为字符串内容() {
        ReportData reportData = new ReportData();
        reportData.setKey("key","value");
        reportData.setKey("key","value1");
        assertThat(reportData.getValue("key"),snapshotMatch(this,"add_strings"));
    }

    @Test
    public void 当设置Key多次自动合并成数组() {
        ReportData reportDataA = new ReportData();
        reportDataA.addKey("name", new StringContent("value1"));
        reportDataA.addKey("name", new StringContent("value2"));
        reportDataA.addKey("name", new StringContent("value3"));
        assertThat(reportDataA.getValue("name"), is(new StringContent(Arrays.asList("value1", "value2", "value3"))));
    }
    @Test
    public void 当设置Key多次自动合并成数组2() {
        ReportData reportDataA = new ReportData();
        reportDataA.addKey("name", "value1");
        reportDataA.addKey("name", "value2");
        reportDataA.addKey("name", "value3");
        assertThat(reportDataA.getValue("name"), is(new StringContent(Arrays.asList("value1", "value2", "value3"))));
    }

    @Test
    public void 当设置图片Key多次自动合并成数组() {
        ReportData reportDataA = new ReportData();
        reportDataA.addKey("name", new ImageContent(Arrays.asList("123".getBytes())));
        reportDataA.addKey("name", new ImageContent(Arrays.asList("456".getBytes())));
        reportDataA.addKey("name", new ImageContent(Arrays.asList("789".getBytes())));
        assertThat(((ImageContent) reportDataA.getValue("name")).size(), is(3));
    }
}
