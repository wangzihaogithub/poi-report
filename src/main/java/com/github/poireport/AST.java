package com.github.poireport;

import java.util.*;
import java.util.function.Predicate;

/**
 * 抽象语法树
 * @author wangzihao 2020-11-10
 */
public interface AST  {
    /**
     * 节点名称
     * @return #text, p, strong, span
     */
    String getName();
    AST getParent();
    List<? extends AST> getChildren();

    boolean hasAttr(String attributeKey);
    String attr(String attributeKey);
    Iterable<? extends Map.Entry<String, String>> attr();

    default String getContent(){
        String content = attr(getName());
        if(content == null || content.isEmpty()){
            StringBuilder builder = new StringBuilder();
            for (AST child : getChildren()) {
                builder.append(child.getContent());
            }
            content = builder.toString();
        }
        return content;
    }
    default boolean isSpan(){
        return Objects.equals("span",getName());
    }
    default boolean isStrong(){
        return Objects.equals("strong",getName());
    }
    default boolean isText(){
        return Objects.equals("#text",getName());
    }

    default Map<String,String> attrMap(String attributeKey){
        String attr = attr(attributeKey);
        Map<String,String> result = new LinkedHashMap<>();
        if(attr != null && attr.length() > 0) {
            for (String s : attr.split(";")) {
                String[] split = s.split(":");
                String key = split[0].trim();
                String humpKey = StringUtil.lineToHump(key);
                if (split.length == 1) {
                    result.put(humpKey, "true");
                } else {
                    result.put(humpKey, split[1].trim());
                }
            }
        }
        return result;
    }

    default List<? extends AST> collect(Predicate<AST> collectChild, Predicate<AST> collect){
        List<AST> result = new ArrayList<>();
        Queue<AST> stack = new LinkedList<>();
        stack.add(this);
        while (!stack.isEmpty()){
            AST top = stack.poll();
            if(collect.test(top)) {
                result.add(top);
            }
            if(collectChild.test(top)) {
                stack.addAll(top.getChildren());
            }
        }
        return result;
    }

    default List<? extends AST> collect(){
        return collect(e->{
            if(e.isStrong() || e.isSpan()){
                return false;
            }
            return true;
        },e-> e.isStrong() || e.isSpan() || e.isText());
    }

    public static void main(String[] args) {
        AST ast6 = JsoupAST.parsePart("<p><span style=\"color: rgb(17, 17, 17);\">5年</span>市场<span style=\"color: rgb(17, 17, 17);\">推广和营销策划经历，4年面向客户和渠道的经验，以客户需求为导向，具有较强</span>市场<span style=\"color: rgb(17, 17, 17);\">敏锐度</span></p>");


        AST ast5 = JsoupAST.parsePart("<span style=\"color: rgb(102, 102, 102)；“>（说个题外话，我今年听到PM和客户聊天的内容）。我们的同志太年轻了，工作3年的，已经思考战略的问题了。每天睡到9点才起床工作，因为晚上还要玩到凌晨一两点睡觉。然后还要抱怨工作压力大。然后还希望给每天6点起床，9点结束工作的，已经工作10年以上的人去做咨询和专家系统呢？不存在的哈。</span>"
        );

        AST ast = JsoupAST.parsePart(
                "候选人<jjj>对头条有</jjj>兴趣 中国地质大学本硕出身，电子信息工程和通信专业双重背景，先后就职于百融金服、迪堡金融、武汉天喻 <span style=\" das; font-size: 5px; color: rgb(255, 0, 0);\">四川省、云南省银行客户资源：如四川建行、四川农行、四川中行、成都银行、成都农商行等客户；</span>  付速度"
        );



        AST ast1 = JsoupAST.parsePart("<span>发</span>");
        assert "发".equals(ast1.getChildren().get(0).getContent());

        AST ast2 = JsoupAST.parsePart("<p>顶顶顶<strong>啊啊啊</strong></p>");
        assert ast2.getChildren().size() == 2;

        AST ast3 = JsoupAST.parsePart("啊<啊>啊");
        assert "啊<啊>啊".equals(ast3.getContent());

        AST ast4 = JsoupAST.parsePart("<span style=\"font-size: 14px; color: rgb(230, 0, 0);\">T输入框</span>舒舒服服<strong>项目描述1</strong>舒舒服服");

        Collection<? extends AST> contents = ast1.collect();
        System.out.println();

    }

}
