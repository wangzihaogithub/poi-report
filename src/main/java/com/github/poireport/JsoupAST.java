package com.github.poireport;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

import java.util.*;

/**
 *
 * {
 *     "p":"<p>顶顶顶<strong>啊啊啊</strong></p>",
 *     "children":[
 *         {
 *             "#text":"顶顶顶"
 *         },
 *         {
 *             "strong":"<strong>啊啊啊</strong>",
 *             "children":[
 *                 {
 *                     "#text":"啊啊啊"
 *                 }
 *             ]
 *         }
 *     ]
 * }
 */
public class JsoupAST implements AST {
    private final Node node;
    private final AST parent;
    private List<JsoupAST> children;

    public JsoupAST(Node node, AST parent) {
        this.node = node;
        if(parent == null && node.parent() != null){
            this.parent = new JsoupAST(node.parent(),null);
        }else {
            this.parent = parent;
        }
    }

    public static AST parsePart(String htmlPart){
        String replace =  htmlPart
                .replace("：",":")
                .replace("“","\"")
                .replace("；",";");
        AST ast = parsePart0(replace);
        if(ast == null){
            throw new IllegalStateException("解析失败 -> "+ htmlPart);
        }
        return ast;
    }

    private static AST parsePart0(String htmlPart){
        Document parse1 = Jsoup.parse(htmlPart);
        Element body = parse1.getElementsByTag("body").get(0);
        Node part;
        if(body.childNodeSize() == 0){
            return null;
        }else if(body.childNodeSize() == 1){
            part = body.childNode(0);
        }else {
            part = body;
        }
        return new JsoupAST(part,null);
    }
    @Override
    public String getName() {
        return node.nodeName();
    }

    @Override
    public AST getParent(){
        return parent;
    }

    @Override
    public String attr(String attributeKey) {
        return node.attr(attributeKey);
    }

    @Override
    public Iterable<? extends Map.Entry<String, String>> attr(){
        return node.attributes();
    }

    @Override
    public boolean hasAttr(String attributeKey) {
        return node.hasAttr(attributeKey);
    }

    @Override
    public List<? extends AST> getChildren() {
        if(children == null) {
            children = new ArrayList<>();
            for (Node node : node.childNodes()) {
                children.add(new JsoupAST(node, this));
            }
        }
        return children;
    }

    @Override
    public String toString() {
        StringJoiner joiner = new StringJoiner(",","{","}");
        joiner.add("\""+getName()+"\":\"" + getContent() + "\"");
        Collection<? extends AST> children = getChildren();
        if(children != null && children.size() > 0){
            joiner.add("\"children\":"+children);
        }
        return joiner.toString();
    }
}
