package com.lozic.genpptx;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystem;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.imaging.ImageInfo;
import org.apache.commons.imaging.ImageReadException;
import org.apache.commons.imaging.Imaging;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * Replaces image placeholders in pptx by images from model.
 *
 */
public class ImageReplacer {

    private static final long INCH_TO_EMU = 914400;

    public State replace(FileSystem fs, Path slideXml, Map<String, Object> model) throws IOException,
            ParserConfigurationException, SAXException, XPathExpressionException, DOMException, ImageReadException {
        InputStream xmlInput = Files.newInputStream(slideXml);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document slideDoc = builder.parse(xmlInput);
        XPath xpath = XPathFactory.newInstance().newXPath();
        XPathExpression expr = xpath.compile("/sld/cSld/spTree/pic/nvPicPr/cNvPr[starts-with(@name, '${')]");
        NodeList nodes = (NodeList) expr.evaluate(slideDoc, XPathConstants.NODESET);

        Path slideXmlRel = fs.getPath("/ppt/slides/_rels/slide1.xml.rels");
        xmlInput = Files.newInputStream(slideXmlRel);
        Document relsDoc = builder.parse(xmlInput);

        for (int i = 0; i < nodes.getLength(); i++) {
            Element node = (Element) nodes.item(i);
            String name = node.getAttribute("name");
            if (name.endsWith("}")) {
                name = name.substring(2, name.length() - 1);
                String imageId = node.getParentNode().getNextSibling().getFirstChild().getAttributes()
                        .getNamedItem("r:embed").getNodeValue();
                expr = xpath.compile("//Relationship[@Id='" + imageId + "']");
                NodeList rels = (NodeList) expr.evaluate(relsDoc, XPathConstants.NODESET);
                String zipPath = rels.item(0).getAttributes().getNamedItem("Target").getNodeValue();
                Path fileInsideZipPath = fs.getPath(zipPath.replaceFirst("\\.\\.", "/ppt"));
                Path imagePath = (Path) model.get(name);
                if (imagePath != null) {
                    Files.copy(imagePath, fileInsideZipPath, StandardCopyOption.REPLACE_EXISTING);
                }
                Element picLocks = (Element) node.getNextSibling().getFirstChild();
                if (!"1".equals(picLocks.getAttribute("noChangeAspect"))) {

                    Element pic = (Element) node.getParentNode().getParentNode();
                    Element xfrm = (Element) pic.getElementsByTagName("p:spPr").item(0).getFirstChild();
                    Element off = (Element) xfrm.getElementsByTagName("a:off").item(0);
                    Element ext = (Element) xfrm.getElementsByTagName("a:ext").item(0);

                    long x = Long.parseLong(off.getAttribute("x"));
                    long y = Long.parseLong(off.getAttribute("y"));

                    long wp = Long.parseLong(ext.getAttribute("cx"));
                    long hp = Long.parseLong(ext.getAttribute("cy"));

                    ImageInfo imageInfo = Imaging.getImageInfo(imagePath.toFile());
                    long wi = Math.round(imageInfo.getPhysicalWidthInch() * INCH_TO_EMU);
                    long hi = Math.round(imageInfo.getPhysicalHeightInch() * INCH_TO_EMU);

                    long w = wp;
                    long h = hp;
                    long dx = 0;
                    long dy = 0;

                    if (wp / hp > wi / hi) {
                        w = wi * hp / hi;
                        dx = (wp - w) / 2;
                    } else if (wp / hp < wi / hi) {
                        h = hi * wp / wi;
                        dy = (hp - h) / 2;
                    }
                    if (wp / hp != wi / hi) {
                        off.setAttribute("x", "" + (x + dx));
                        off.setAttribute("y", "" + (y + dy));
                        ext.setAttribute("cx", "" + w);
                        ext.setAttribute("cy", "" + h);
                    }
                }
            }
        }
        return new State(slideDoc, relsDoc);
    }
}
