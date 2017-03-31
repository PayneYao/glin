package com.lozic.genpptx;

import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

public class PptxGeneratorTest {

    private Path templatePath;
    private Path cssPath;
    private Path outputPath;
    private Path out5;
    private PptxGenerator generator = PptxGeneratorFactory.newPptxGenerator();

    @Before
    public void init() throws URISyntaxException, IOException {
        templatePath = Paths.get(this.getClass().getResource("test.pptx").toURI());
        out5 =  Paths.get(this.getClass().getResource("out5.pptx").toURI());
        cssPath = Paths.get(this.getClass().getResource("test.css").toURI());
        outputPath = Paths.get("target/out.pptx");
        Files.deleteIfExists(outputPath);
    }

    private void process(Map<String, Object> model) {

        try {
            generator.process(templatePath, cssPath, outputPath, model);
        } catch (IOException | GenerationException e) {
            Assert.fail(e.getMessage());
        }
        Assert.assertTrue(Files.exists(outputPath));
    }

    @Test
    public void testCreatePPT2007Slide(){

    }

    @Test
    public void testProcessSimpleText() {
        Map<String, Object> model = new HashMap<>();
        model.put("testText", "this is test");
        process(model);
    }

    @Test
    public void testProcessImage() throws URISyntaxException {
        Map<String, Object> model = new HashMap<>();
        model.put("testImage", Paths.get(this.getClass().getResource("signals.png").toURI()));
        process(model);
    }

    @Test
    public void testProcessHtml() throws URISyntaxException {
        Map<String, Object> model = new HashMap<>();
        model.put("testHtml",
                "simple text <br/>" + "<b>bold</b> <i>italic</i> <u>underlined</u></br>"
                        + "<span style='color: rgb(255, 0, 0);'>I'm red "+Paths.get(this.getClass().getResource("signals.png").toURI())+"</span>"
                        + "<img src='"+Paths.get(this.getClass().getResource("signals.png").toURI())+"'>"
                        + "<span style='font-size: 60pt;'>I'ma big,</span> <span style='font-size: 9pt;'>me-tiny,</span> I am standard<br/>"
                        + "<a href='http://www.google.com'>Search on Google</a>");
        process(model);
    }

    @Test
    public void testProcessCss() {
        Map<String, Object> model = new HashMap<>();
        model.put("testHtml", "<span class='test-class'>I'm styled in css by classname.</span>");
        process(model);
    }

    @Test
    public void testProcessDelete() {
        Map<String, Object> model = new HashMap<>();
        model.put("testDelete", null);
        process(model);
    }
}
