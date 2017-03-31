package com.lozic.genpptx;

import com.lozic.genpptx.css.CssInline;
import com.lozic.genpptx.html.ASupport;
import com.lozic.genpptx.html.BSupport;
import com.lozic.genpptx.html.BrSupport;
import com.lozic.genpptx.html.Html2PptxTransformer;
import com.lozic.genpptx.html.ISupport;
import com.lozic.genpptx.html.LiSupport;
import com.lozic.genpptx.html.TextSupport;
import com.lozic.genpptx.html.Transformer;
import com.lozic.genpptx.html.USupport;
import com.lozic.genpptx.html.UlSupport;

public class PptxGeneratorFactory {

    public static PptxGenerator newPptxGenerator() {
        CssInline cssInline = new CssInline();
        Transformer transformer = new Html2PptxTransformer(cssInline);
        transformer.setSupportSet(new ASupport(), new BrSupport(), new BSupport(transformer), new ISupport(transformer),
                new LiSupport(transformer), new TextSupport(), new UlSupport(transformer), new USupport(transformer));
        VariableProcessor variableProcessor = new VariableProcessor(transformer);
        ImageReplacer replacer = new ImageReplacer();
        PptxGenerator generator = new PptxGenerator(variableProcessor, replacer, new Deletor());
        return generator;
    }

}
