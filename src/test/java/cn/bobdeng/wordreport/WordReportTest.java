package cn.bobdeng.wordreport;

import cn.bobdeng.testtools.TestResource;
import com.google.common.io.Resources;
import org.junit.After;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Base64;
import java.util.List;
import static cn.bobdeng.wordreport.Placeholder.PLACEHOLDER_BEGIN;
import static cn.bobdeng.wordreport.Placeholder.PLACEHOLDER_END;
import static org.hamcrest.core.Is.is;
import static org.junit.Assert.*;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertThat;
import static org.junit.Assert.assertTrue;


public class WordReportTest {
    private String fileName = "output.docx";

    @After
    public void setup() {
        File file = new File(fileName);
        if (file.exists()) {
            file.delete();
        }
    }

    @Test
    public void replace_single_text() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("检测人", new StringContent("邓志国"));
        String resourceName = "doc1.docx";
        byte[] output = exportReport(reportData, resourceName);
        fileName = "a1.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("邓志国"));

    }


    @Test
    public void Given报告表格数据和带表格块带模板_When输出报告_Then输出表格() throws Exception {
        ReportData reportData = new ReportData();
        List<String> titles = Arrays.asList("样品编号", "采样地点");
        List<String> line = Arrays.asList("0001", "采样地点1");
        List<List<String>> lines = Arrays.asList(line);
        TemplateContent tableContent = new TableContent(titles, lines);
        reportData.setKey("表格块名称", tableContent);
        reportData.setKey("样品检测结果", tableContent);
        byte[] templateFile = new TestResource(this, "表格块.docx").readBytes();
        new TestResource(this, fileName).save(reportData.output(templateFile));
        WordTextReader reader = new WordTextReader(new TestResource(this, fileName).getFile());
        assertTrue(reader.getText().contains("样品编号"));
        assertTrue(reader.getText().contains("0001"));

    }

    private byte[] exportReport(ReportData reportData, String resourceName) throws IOException {
        byte[] templateFile = Resources.toByteArray(Resources.getResource(resourceName));
        return reportData.output(templateFile);
    }

    @Test
    public void replace_single_text_文字不在一个段落里() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("检测记录", new StringContent("邓志国"));
        byte[] output = exportReport(reportData, "doc2.docx");
        fileName = "a12.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("邓志国"));
    }


    @Test
    public void replace_single_03() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("姓名", new StringContent("邓志国"));
        reportData.setKey("任务编号", new StringContent("100-100"));
        byte[] output = exportReport(reportData, "03.docx");
        fileName = "temp_out.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("邓志国"));
        assertFalse(reader.getText().contains("100-100"));

    }

    @Test
    public void replace_array_03() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("检验项目[]", new StringContent(Arrays.asList("项目1", "项目2")));
        byte[] output = exportReport(reportData, "array.docx");
        fileName = "temp_out.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("项目1"));
        assertTrue(reader.getText().contains("项目2"));

    }

    @Test
    public void 数组单个混合编排() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("上机检测.样品编号数据[]", new StringContent(Arrays.asList("项目1", "项目2")));
        reportData.setKey("委托单位数据", new StringContent("广东"));
        byte[] output = exportReport(reportData, "检测报告最新模板-3.docx");
        fileName = "检测报告最新模板-3-1.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("项目1"));
        assertTrue(reader.getText().contains("项目2"));
        assertTrue(reader.getText().contains("广东"));
    }

    @Test
    public void 数组内容是空() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("委托单位数据", new StringContent("广东"));
        byte[] output = exportReport(reportData, "检测报告最新模板-3.docx");
        fileName = "temp_out.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("广东"));
    }

    @Test
    public void 加数组内容() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("任务名称", new StringContent("任务名称123"));
        reportData.setKey("考核序号[]", new StringContent(Arrays.asList("考核序号1", "考核序号2")));
        reportData.setKey("要素条款[]", new StringContent(Arrays.asList("")));
        reportData.setKey("要素类型[]", new StringContent(Arrays.asList("组织", "人员", "其他")));
        reportData.setKey("不符合类型[]", new StringContent(Arrays.asList("11", "22")));
        byte[] output = exportReport(reportData, "不符合项汇总表1.docx");
        fileName = "不符合项汇总表1.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("任务名称123"));
        assertTrue(reader.getText().contains("考核序号1"));
        assertTrue(reader.getText().contains("考核序号2"));
    }

    @Test
    public void 一个单元格有两个占位符() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("任务编号", new StringContent(Arrays.asList("xxx1234")));
        reportData.setKey("考核序号[]", new StringContent(Arrays.asList("1", "2")));
        reportData.setKey("要素类型[]", new StringContent(Arrays.asList("组织", "人员", "其他")));
        fileName = "两个占位符.docx";
        byte[] output = exportReport(reportData, fileName);
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("1、组织"));
        assertTrue(reader.getText().contains("2、人员"));
        assertTrue(reader.getText().contains("、其他"));
    }

    @Test
    public void 页眉测试() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("受理.编号", new StringContent(Arrays.asList("xxx1234")));
        fileName = "测试页眉.docx";
        byte[] output = exportReport(reportData, fileName);
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("xxx1234"));

    }

    @Test
    public void 替换页脚占位符() throws Exception {
        ReportData reportData = new ReportData();
        reportData.setKey("页脚占位符", new StringContent(Arrays.asList("xxx1234")));
        reportData.setKey("签名", new ImageContent(new TestResource(this, "sign.jpg").readBytes()));
        WordTextReader reader = getWordTextReader(reportData, "页脚有占位符.docx");
        assertTrue(reader.getText().contains("xxx1234"));
        assertTrue(reader.getText().contains("image:/9j/4AAQSkZJ"));
    }

    @Test
    public void 替换页脚占位符签名() throws Exception {
        ReportData reportData = new ReportData();
        reportData.addKey("检测者", new ImageContent(new TestResource(this, "sign.jpg").readBytes()));
        reportData.addKey("检测者", new ImageContent(new TestResource(this, "sign1.jpg").readBytes()));
        WordTextReader reader = getWordTextReader(reportData, "页脚有签名1.docx");
        assertTrue(reader.getText().contains("image:/9j/4AAQSkZJ"));
    }

    @Test
    public void Given占位符是空_When生成文档_Then生成空字符串() throws Exception {
        ReportData reportData = new ReportData();
        String content = null;
        reportData.setKey("页脚占位符", content);
        reportData.setKey("页眉占位符", content);
        WordTextReader reader = getWordTextReader(reportData, "页脚有占位符.docx");
        assertFalse(reader.getText().contains("null"));
    }

    private WordTextReader getWordTextReader(ReportData reportData, String templateResourceFile) throws Exception {
        fileName = "output.docx";
        byte[] templateFile = new TestResource(this, templateResourceFile).readBytes();
        new TestResource(this, fileName).save(reportData.output(templateFile));
        return new WordTextReader(new TestResource(this, fileName).getFile());
    }

    @Test
    public void Given占位符用新行_When生成文档_Then多行() throws Exception {
        ReportData reportData = new ReportData();
        reportData.addKey("占位符", "hello");
        reportData.addKey("占位符", "world");
        WordTextReader wordTextReader = getWordTextReader(reportData, "占位符带新行标记.docx");
        assertFalse(wordTextReader.getText().contains("hello world"));
        assertTrue(wordTextReader.getText().contains("hello"));
        assertTrue(wordTextReader.getText().contains("world"));
    }
    @Test
    public void Given占位符用唯一标记_When生成文档_Then去掉重复() throws Exception {
        ReportData reportData = new ReportData();
        reportData.addKey("总结论.检验标准", "hello");
        reportData.addKey("总结论.检验标准", "hello");
        WordTextReader wordTextReader = getWordTextReader(reportData, "公共场所-评价报告模板.docx");
        assertFalse(wordTextReader.getText().contains("hello hello"));
        assertTrue(wordTextReader.getText().contains("hello"));
    }

    @Test
    public void 一个单元格有两个占位符2() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("上机检测.检测标准[]", new StringContent(Arrays.asList("标准1", "标准2")));
        reportData.setKey("上机检测.检测方法[]", new StringContent(Arrays.asList("方法1", "方法2")));
        fileName = "test11.docx";
        byte[] output = exportReport(reportData, fileName);
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("标准1方法1"));
        assertTrue(reader.getText().contains("标准2方法2"));
    }

    @Test
    public void 当单个值不存在但存在数组则取数组值() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("任务名称", new StringContent("任务名称"));
        reportData.setKey("考核序号[]", new StringContent(Arrays.asList("1", "2")));
        reportData.setKey("要素条款[]", new StringContent(Arrays.asList("22", "223")));
        reportData.setKey("要素类型[]", new StringContent(Arrays.asList("组织", "人员")));
        reportData.setKey("不符合类型[]", new StringContent(Arrays.asList("11", "22")));
        assertThat(reportData.getContent(PLACEHOLDER_BEGIN + "考核序号" + PLACEHOLDER_END), is(new StringContent(Arrays.asList("1", "2"))));
        assertThat(reportData.getContent(PLACEHOLDER_BEGIN + "要素条款" + PLACEHOLDER_END), is(new StringContent(Arrays.asList("22", "223"))));
    }

    @Test
    public void 替换Word里面自定义表格内容() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("任务名称", new StringContent("任务名称123"));
        reportData.setKey("考核序号[]", new StringContent(Arrays.asList("序号1", "序号2", "序号3")));
        reportData.setKey("要素条款[]", new StringContent(Arrays.asList("22")));
        reportData.setKey("要素类型[]", new StringContent(Arrays.asList("组织", "人员")));
        reportData.setKey("不符合类型[]", new StringContent(Arrays.asList("11", "22")));
        fileName = "带自定义表格.docx";
        byte[] output = exportReport(reportData, fileName);
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("任务名称123"));
        assertTrue(reader.getText().contains("序号1"));
        assertTrue(reader.getText().contains("序号3"));

    }

    @Test
    public void 占位符黏连的情况() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("考核序号", Arrays.asList("序号1", "序号2", "序号3"));
        reportData.setKey("要素类型", new StringContent(Arrays.asList("组织", "人员")));
        reportData.setKey("检测标准[]", new StringContent(Arrays.asList("标准1", "标准2")));
        reportData.setKey("检测方法[]", new StringContent(Arrays.asList("方法1", "方法2")));
        fileName = "占位符黏连.docx";
        byte[] output = exportReport(reportData, fileName);
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("序号1"));
        assertTrue(reader.getText().contains("序号2"));
        assertTrue(reader.getText().contains("序号3"));
        assertTrue(reader.getText().contains("组织"));
        assertTrue(reader.getText().contains("人员"));

    }

    @Test
    public void 替换页眉里面的值() throws IOException {
        ReportData reportData = new ReportData();
        reportData.setKey("任务单.委托单号数据", new StringContent("123456"));
        byte[] output = exportReport(reportData, "value_in_header.docx");
        fileName = "value_in_header_out.docx";
        new FileOutputStream(fileName).write(output);
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains("（123456）"));
    }

    @Test
    public void 插入图片签名() throws IOException {
        ReportData reportData = new ReportData();
        byte[] image1 = Resources.toByteArray(Resources.getResource("sign.jpg"));
        reportData.addKey("操作人签名", new ImageContent(image1));
        byte[] image2 = Resources.toByteArray(Resources.getResource("sign1.jpg"));
        reportData.addKey("检验人签名", new ImageContent(image2));
        byte[] templateFile = Resources.toByteArray(Resources.getResource("sign.docx"));
        fileName = "图片签名.docx";
        new FileOutputStream(fileName).write(reportData.output(templateFile));
        WordTextReader reader = new WordTextReader(new File(fileName));
        assertTrue(reader.getText().contains(Base64.getEncoder().encodeToString(image1)));
        assertTrue(reader.getText().contains(Base64.getEncoder().encodeToString(image2)));
    }

    @Test
    public void 替换图片签名在表格里() throws Exception {
        ReportData reportData = new ReportData();
        byte[] image1 = new TestResource(this, "sign.jpg").readBytes();
        reportData.addKey("操作人签名", new ImageContent(image1));
        byte[] image2 = new TestResource(this, "sign1.jpg").readBytes();
        reportData.addKey("检验人签名", new ImageContent(image2));
        byte[] templateFile = new TestResource(this, "sign_in_table.docx").readBytes();
        fileName = "图片签名在表格里.docx";
        new TestResource(this, fileName).save(reportData.output(templateFile));
        WordTextReader reader = new WordTextReader(new TestResource(this, fileName).getFile());
        assertTrue(reader.getText().contains(Base64.getEncoder().encodeToString(image1)));
        assertTrue(reader.getText().contains(Base64.getEncoder().encodeToString(image2)));
    }

    @Test
    public void 插入图片签名在表格里() throws Exception {
        ReportData reportData = new ReportData();
        reportData.addKey("操作人签名[]", new ImageContent(new TestResource(this, "sign.jpg").readBytes()));
        reportData.addKey("操作人签名[]", new ImageContent(new TestResource(this, "sign.jpg").readBytes()));
        reportData.addKey("检验人签名[]", new ImageContent(new TestResource(this, "sign1.jpg").readBytes()));
        byte[] templateFile = new TestResource(this, "sign_in_table2.docx").readBytes();
        fileName = "图片签名在表格里2.docx";
        new TestResource(this, fileName).save(reportData.output(templateFile));
        WordTextReader reader = new WordTextReader(new TestResource(this, fileName).getFile());
        assertTrue(reader.getText().contains(Base64.getEncoder().encodeToString(new TestResource(this, "sign.jpg").readBytes())));
        assertTrue(reader.getText().contains(Base64.getEncoder().encodeToString(new TestResource(this, "sign1.jpg").readBytes())));
    }

}
