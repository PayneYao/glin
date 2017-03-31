package com.lozic.genpptx.css;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;

public class CssInline {

    public void applyCss(CSSStyleSheet css, org.jsoup.nodes.Document html) {
        Map<Element, Map<String, String>> styles = traverse(css, html);
        apply(styles);
    }

    private void apply(Map<Element, Map<String, String>> styles) {
        for (Entry<Element, Map<String, String>> style : styles.entrySet()) {
            Element element = style.getKey();
            Map<String, String> map = style.getValue();
            StringBuilder builder = new StringBuilder();
            for (Entry<String, String> css : map.entrySet()) {
                builder.append(css.getKey()).append(":").append(css.getValue()).append(";");
            }
            builder.append(element.attr("style"));
            element.attr("style", builder.toString());
            element.removeAttr("class");
        }
    }

    private Map<Element, Map<String, String>> traverse(CSSStyleSheet css, org.jsoup.nodes.Document html) {
        CSSRuleList rules = css.getCssRules();
        Map<Element, Map<String, String>> styles = new HashMap<>();

        for (int i = 0; i < rules.getLength(); i++) {
            CSSRule rule = rules.item(i);
            if (rule instanceof CSSStyleRule) {
                CSSStyleRule styleRule = (CSSStyleRule) rule;
                String selector = styleRule.getSelectorText();

                if (!selector.contains(":")) {
                    traverseSelected(html, styles, styleRule, selector);
                }
            }
        }
        return styles;
    }

    private void traverseSelected(org.jsoup.nodes.Document html, Map<Element, Map<String, String>> styles,
            CSSStyleRule styleRule, String selector) {
        final Elements selectedElements = html.select(selector);
        for (Element selected : selectedElements) {
            traverseElement(styles, styleRule, selected);

        }
    }

    private void traverseElement(Map<Element, Map<String, String>> styles, CSSStyleRule styleRule, Element element) {
        if (!styles.containsKey(element)) {
            styles.put(element, new LinkedHashMap<String, String>());
        }

        final CSSStyleDeclaration styleDeclaration = styleRule.getStyle();

        for (int j = 0; j < styleDeclaration.getLength(); j++) {
            final String propertyName = styleDeclaration.item(j);
            final String propertyValue = styleDeclaration.getPropertyValue(propertyName);
            final Map<String, String> elementStyle = styles.get(element);
            elementStyle.put(propertyName, propertyValue);
        }
    }
}