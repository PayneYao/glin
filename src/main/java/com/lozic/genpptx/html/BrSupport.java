package com.lozic.genpptx.html;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.w3c.dom.Document;

import com.lozic.genpptx.GenerationException;
import com.lozic.genpptx.State;

public class BrSupport implements NodeSupport {

    @Override
    public boolean supports(Node node) {
        if (node instanceof Element) {
            return "br".equals(((Element) node).tagName());
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        org.w3c.dom.Element p = state.getP();
        Document slideDoc = state.getSlideDoc();
        org.w3c.dom.Element br = slideDoc.createElement("a:br");
        p.appendChild(br);
        org.w3c.dom.Element rPr = slideDoc.createElement("a:rPr");
        br.appendChild(rPr);
        rPr.setAttribute("lang", "en-US");
        rPr.setAttribute("dirty", "0");
        rPr.setAttribute("smtClean", "0");
    }

}
