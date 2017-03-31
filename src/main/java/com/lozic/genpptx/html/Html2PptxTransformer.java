package com.lozic.genpptx.html;

import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

import org.apache.commons.beanutils.BeanUtils;
import org.jsoup.Jsoup;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.css.CSSStyleSheet;

import com.lozic.genpptx.GenerationException;
import com.lozic.genpptx.State;
import com.lozic.genpptx.css.CssInline;
import com.lozic.genpptx.css.CssProcessor;
import com.lozic.genpptx.css.Style;

public class Html2PptxTransformer implements Transformer {

    private Set<NodeSupport> supportSet;

    private CssInline cssInline;

    public Html2PptxTransformer(CssInline cssInline) {
        super();
        this.cssInline = cssInline;
    }

    public void setSupportSet(NodeSupport... supports) {
        this.supportSet = new HashSet<>(Arrays.asList(supports));
    }

    public void convert(State state, CSSStyleSheet css, String htmlString) throws GenerationException {
        org.jsoup.nodes.Document html = Jsoup.parse(htmlString);
        cssInline.applyCss(css, html);
        org.jsoup.nodes.Node body = html.body();

        Document slideDoc = state.getSlideDoc();
        Element p = slideDoc.createElement("a:p");
        state.setP(p);
        state.getTxBody().appendChild(p);
        Element pPr = slideDoc.createElement("a:pPr");
        p.appendChild(pPr);
        pPr.setAttribute("algn", "l");
        iterate(state, body);
    }

    public void iterate(State state, org.jsoup.nodes.Node node) throws GenerationException {
        try {
            for (org.jsoup.nodes.Node htmlNode : node.childNodes()) {
                traverse(state, htmlNode);
            }
        } catch (Exception e) {
            throw new GenerationException(e);
        }
    }

    private void traverse(State state, org.jsoup.nodes.Node node) throws GenerationException {
        try {
            Style style = state.getStyle();
            Style prevStyle = (Style) BeanUtils.cloneBean(style);
            String htmlStyle = node.attr("style");
            CssProcessor.process(htmlStyle, style);

            NodeSupport support = null;
            for (NodeSupport sup : supportSet) {
                if (sup.supports(node)) {
                    support = sup;
                    break;
                }
            }
            if (support != null) {
                support.process(state, node);
            } else {
                iterate(state, node);
            }
            state.setStyle(prevStyle);
        } catch (GenerationException e) {
            throw e;
        } catch (IllegalAccessException | InstantiationException | InvocationTargetException
                | NoSuchMethodException e) {
            throw new GenerationException(e);
        }
    }

}
