package com.lozic.genpptx.html;

import org.jsoup.nodes.Node;

import com.lozic.genpptx.GenerationException;
import com.lozic.genpptx.State;

public interface NodeSupport {

    boolean supports(Node node);

    void process(State state, Node node) throws GenerationException;

}
