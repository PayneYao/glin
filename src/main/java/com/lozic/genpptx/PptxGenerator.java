package com.lozic.genpptx;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.css.sac.InputSource;
import org.w3c.dom.css.CSSStyleSheet;

import com.steadystate.css.parser.CSSOMParser;

/**
 * Class for generating pptx files.
 *
 */

public class PptxGenerator {

    private VariableProcessor variableProcessor;

    private ImageReplacer replacer;

    private Deletor deletor;

    public PptxGenerator(VariableProcessor variableProcessor, ImageReplacer replacer, Deletor deletor) {
        super();
        this.variableProcessor = variableProcessor;
        this.replacer = replacer;
        this.deletor = deletor;
    }

    public void process(Path templatePath, Path cssPath, Path outputPath, Map<String, Object> model)
            throws IOException, GenerationException {
        try {
            Files.copy(templatePath, outputPath, StandardCopyOption.REPLACE_EXISTING);
            FileSystem fs = FileSystems.newFileSystem(outputPath, null);

            Path slideXml = fs.getPath("/ppt/slides/slide1.xml");
            Path relXml = fs.getPath("/ppt/slides/_rels/slide1.xml.rels");
            State state = replacer.replace(fs, slideXml, model);

            deletor.process(state, model);

            CSSOMParser parser = new CSSOMParser();
            CSSStyleSheet css = parser.parseStyleSheet(
                    new InputSource(Files.newBufferedReader(cssPath, StandardCharsets.UTF_8)), null, null);

            variableProcessor.process(state, css, model);
            savePptx(slideXml, relXml, state);

            fs.close();
        } catch (Exception e) {
            throw new GenerationException("Could not generate resulting ppt.", e);
        }
    }

    private void savePptx(Path slideXml, Path relXml, State state) throws TransformerConfigurationException,
            TransformerFactoryConfigurationError, IOException, TransformerException {

        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        Path outSlideXml = Files.createTempFile(null, ".xml");
        Result output = new StreamResult(Files.newOutputStream(outSlideXml));
        Source input = new DOMSource(state.getSlideDoc());
        transformer.transform(input, output);

        Files.copy(outSlideXml, slideXml, StandardCopyOption.REPLACE_EXISTING);
        Files.delete(outSlideXml);

        transformer = TransformerFactory.newInstance().newTransformer();
        outSlideXml = Files.createTempFile(null, ".xml");
        output = new StreamResult(Files.newOutputStream(outSlideXml));
        input = new DOMSource(state.getRelDoc());
        transformer.transform(input, output);

        Files.copy(outSlideXml, relXml, StandardCopyOption.REPLACE_EXISTING);
        Files.delete(outSlideXml);

    }

}
