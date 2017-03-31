package com.lozic.genpptx.html;

import org.jsoup.nodes.Node;
import org.w3c.dom.css.CSSStyleSheet;

import com.lozic.genpptx.GenerationException;
import com.lozic.genpptx.State;

public interface Transformer {

    public void setSupportSet(NodeSupport... supports);

    void iterate(State state, Node node) throws GenerationException;

    void convert(State state, CSSStyleSheet css, String htmlString) throws GenerationException;
}
