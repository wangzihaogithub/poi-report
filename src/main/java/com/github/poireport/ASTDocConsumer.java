package com.github.poireport;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import lombok.Getter;
import lombok.Setter;

import java.util.*;

import static com.github.poireport.DeepooveUtil.extendStyle;
import static com.github.poireport.StringUtil.parseNumber;

/**
 * 语法树与doc的配置
 * @author wangzihao 2020年11月10日15:03:59
 */
public class ASTDocConsumer implements java.util.function.Consumer<DeepooveUtil.DocElement> {
    public static final String ATTR_STYLE = "style";
    @Getter@Setter
    private String font = "微软雅黑";
    @Getter
    private final Style defaultStyle = new Style();{
        defaultStyle.setFontSize(10);
        defaultStyle.setFontFamily(font);
    }

    @Override
    public void accept(DeepooveUtil.DocElement element) {
        //语法树
        AST ast = element.getNode();
        //样式继承
        Style style = new Style();
        extendStyle(defaultStyle,style);
        extendStyle(element.root().getStyle(),style);

        //多级需要将写数据的指针右移
        TextRenderData data = element.getData();
        DeepooveUtil.DocElement.RunWrapper dataNode;
        if (data instanceof DeepooveUtil.MultistageTextRenderFilterData) {
            dataNode = getMultistageDataNode(element,style, (DeepooveUtil.MultistageTextRenderFilterData) data);
        }else {
            //非多级在当前指针写数据
            dataNode = element.root();
        }

        //如果需要覆盖一级项目符号
        if(data instanceof DeepooveUtil.TextRenderFilterData) {
            if (((DeepooveUtil.TextRenderFilterData) data).getFirstSymbol() != null) {
                element.setFirstSymbolItem(((DeepooveUtil.TextRenderFilterData) data).getFirstSymbol());
            }
        }

        //数据 样式
        setContentAndStyle(dataNode,ast,style);

        //广度遍历
//        bfs(element,ast.getChildren(),dataNode.getStyle());
    }

    private void bfs(DeepooveUtil.DocElement element, List<? extends AST> child, Style style){
        if(child.isEmpty()){
            return;
        }
        Style currStyle = style;
        Queue<AST> stack = new LinkedList<>(child);
        while (!stack.isEmpty()){
            AST top = stack.poll();

            DeepooveUtil.DocElement.RunWrapper dataNode = element.createRun();
            setContentAndStyle(dataNode,top,currStyle);
            currStyle = dataNode.getStyle();

            stack.addAll(top.getChildren());
        }
    }

    private void setContentAndStyle(DeepooveUtil.DocElement.RunWrapper dataNode, AST top, Style style) {
        //数据
        String content = top.getContent();
        dataNode.setContent(content);

        //样式
        Style rewriteStyle = getStyle(top, style,font);
        dataNode.setStyle(rewriteStyle);
    }

    private DeepooveUtil.DocElement.RunWrapper getMultistageDataNode(DeepooveUtil.DocElement element,Style parentStyle,
                                                                     DeepooveUtil.MultistageTextRenderFilterData data){
        DeepooveUtil.DocElement.RunWrapper dataNode = element.root();
        Map<Integer, String> levelMap = data.getLevelMap();
        if(levelMap == null){
            levelMap = Collections.emptyMap();
        }

        //全部缩进
        int nodeIndex = element.getNodeIndex();
        int level = data.getLevel();
        String levelSymbol = levelMap.getOrDefault(level, "");
        if(nodeIndex == 0) {
            for (int i = 0; i < level; i++) {
                dataNode = element.createRun();
            }
        }
        for (DeepooveUtil.DocElement.RunWrapper wrapper : element) {
            if(wrapper == dataNode){
                continue;
            }
            wrapper.setContent("  ");
            wrapper.symbol();
        }

        //末尾加项目符号
        DeepooveUtil.DocElement.RunWrapper dataNodePrev = dataNode.prev();
        if(dataNodePrev != null){
            dataNodePrev.symbol();
            dataNodePrev.setContent(levelSymbol);
        }
        return dataNode;
    }

    private static Style getStyle(AST ast, Style parentStyle, String font){
        Style currentStyle = parentStyle;
        //加粗
        if(ast.isStrong()){
            currentStyle = onStrong(parentStyle,font);
        }
        //html样式映射到doc中
//        if (ast.hasAttr(ATTR_STYLE)) {
//            Map<String, String> attrMap = ast.attrMap(ATTR_STYLE);
//            currentStyle = onStyle(currentStyle, attrMap);
//        }
        return currentStyle;
    }

    private static Style onStrong(Style parentStyle,String font){
        Style style = new Style();
        extendStyle(parentStyle,style);
        style.setFontFamily(font);
        style.setBold(true);
        return style;
    }

    private Style onStyle(Style parentStyle,Map<String,String> styleAttrMap){
        Style style = new Style();
        extendStyle(parentStyle,style);
        style.setFontFamily(font);
        String fontSize = styleAttrMap.get("fontSize");
        Integer[] fontSizes = parseNumber(fontSize);
        if(fontSizes.length > 0) {
            style.setFontSize(fontSizes[0]);
        }
        String color = styleAttrMap.get("color");
        if(color != null && color.length() > 0) {
            style.setColor(parseColor16(color));
        }
        return style;
    }

    private static String parseColor16(String rgbOr16){
        String color16;
        if(rgbOr16.startsWith("rgb")){
            Integer[] rgbs = parseNumber(rgbOr16);
            if(rgbs.length == 3) {
                color16 = String.format("%02x%02x%02x", rgbs[0], rgbs[1], rgbs[2]).toUpperCase();
            }else {
                color16 = null;
            }
        }else if(rgbOr16.startsWith("#")){
            color16 = rgbOr16.substring(1);
        }else {
            color16 = rgbOr16;
        }
        return color16;
    }

    public static void main(String[] args) {
        Integer[] integer = parseNumber("15px");
        Integer[] px15 = parseNumber("px15");
        Integer[] integers = parseNumber("rgb(230, 0, 0)");

        AST ast6 = JsoupAST.parsePart("<p><span style=\"color: rgb(17, 17, 17);\">5年</span>市场<span style=\"color: rgb(17, 17, 17);\">推广和营销策划经历，4年面向客户和渠道的经验，以客户需求为导向，具有较强</span>市场<span style=\"color: rgb(17, 17, 17);\">敏锐度</span></p>");

        assert "5年市场推广和营销策划经历，4年面向客户和渠道的经验，以客户需求为导向，具有较强市场敏锐度"
                .equals(ast6.getContent());
        StringBuilder builder = new StringBuilder();
        for (AST top : ast6.getChildren()) {
            builder.append(top.getContent());
        }
        assert "5年市场推广和营销策划经历，4年面向客户和渠道的经验，以客户需求为导向，具有较强市场敏锐度"
                .equals(builder.toString());

        System.out.println();
    }
}
