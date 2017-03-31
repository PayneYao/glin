package com.lozic.genpptx.html;

import org.apache.commons.lang3.StringUtils;
import org.jsoup.nodes.Node;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.lozic.genpptx.GenerationException;
import com.lozic.genpptx.State;
import com.lozic.genpptx.css.Style;

public class LiSupport implements NodeSupport {

    private Transformer transformer;

    public LiSupport(Transformer transformer) {
        super();
        this.transformer = transformer;
    }

    @Override
    public boolean supports(Node node) {
        if (node instanceof org.jsoup.nodes.Element) {
            return "li".equals(((org.jsoup.nodes.Element) node).tagName());
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        Element p = state.getP();
        Style style = state.getStyle();
        Document slideDoc = state.getSlideDoc();
        Element txBody = state.getTxBody();
        if (node.parent().previousSibling() == null && node.previousSibling() == null) {
            p.removeChild(p.getFirstChild());
        }
        Element pPr = slideDoc.createElement("a:pPr");
        p.appendChild(pPr);
        pPr.setAttribute("algn", "l");
        pPr.setAttribute("marL", "342900");
        pPr.setAttribute("indent", "-342900");

        if (StringUtils.isNotBlank(style.getLiColor())) {
            Element buClr = slideDoc.createElement("a:buClr");
            pPr.appendChild(buClr);
            Element srgbClr = slideDoc.createElement("a:srgbClr");
            buClr.appendChild(srgbClr);
            srgbClr.setAttribute("val", style.getLiColor());
        }

        Element buFont = slideDoc.createElement("a:buFont");
        pPr.appendChild(buFont);
        buFont.setAttribute("typeface", "Arial");
        buFont.setAttribute("pitchFamily", "34");
        buFont.setAttribute("charset", "0");

        Element buChar = slideDoc.createElement("a:buChar");
        pPr.appendChild(buChar);
        String bulletChar = StringUtils.isNotBlank(style.getLiChar()) ? style.getLiChar() : "\u2022";
        buChar.setAttribute("char", bulletChar);

        transformer.iterate(state, node);
        p = slideDoc.createElement("a:p");
        txBody.appendChild(p);
        state.setP(p);
        if (node.nextSibling() == null) {
            pPr = slideDoc.createElement("a:pPr");
            p.appendChild(pPr);
            pPr.setAttribute("algn", "l");
        }
    }

}
