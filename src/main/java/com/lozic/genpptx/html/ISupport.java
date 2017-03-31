package com.lozic.genpptx.html;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

import com.lozic.genpptx.GenerationException;
import com.lozic.genpptx.State;

public class ISupport implements NodeSupport {

    private Transformer transformer;

    public ISupport(Transformer transformer) {
        super();
        this.transformer = transformer;
    }

    @Override
    public boolean supports(Node node) {
        if (node instanceof Element) {
            String tagName = ((Element) node).tagName();
            return "i".equals(tagName) || "em".equals(tagName);
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        state.getStyle().setItalic(true);
        transformer.iterate(state, node);
    }

}
