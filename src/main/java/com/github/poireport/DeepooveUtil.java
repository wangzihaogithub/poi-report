package com.github.poireport;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.DocxRenderPolicy;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.StyleUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;

import java.io.ByteArrayInputStream;
import java.io.PrintStream;
import java.lang.reflect.*;
import java.util.*;
import java.util.concurrent.Callable;
import java.util.function.*;

/**
 * 参考资料：1、官方文档：http://deepoove.com/poi-tl/
 * 开源：https://github.com/Sayi/poi-tl/wiki
 * 扩展了自定义字段
 *
 * @author acer01
 */
@Slf4j
public class DeepooveUtil {
    /**DeepooveUtil
     * 为了实现在段落的任意点插入表格的功能, 原生只支持在段落前.
     */
    public static final ThreadLocal<Context> CONTEXT_LOCAL = ThreadLocal.withInitial(Context::new);
    /**
     * 项目符号 样式.
     */
    public static final Style LEVEL_SYMBOL_STYLE = new Style();
    /**
     * 项目符号 不改变模板的设置.
     */
    public static final String LEVEL_SYMBOL_CLOSE = null;
    /**
     * 圆圈  项目符号 2级
     */
    public static final String LEVEL_SYMBOL_YUAN = String.valueOf((char) 0XF050);
    /**
     * 对号 √. 项目符号 2级
     */
    public static final String LEVEL_SYMBOL_DEFAULT = String.valueOf((char) 0XF050);
    /**
     * '❖' (0x2756) 项目符号 3级
     */
    public static final String LEVEL_SYMBOL_3 = String.valueOf((char) 0x2756);
    /**
     * '◊' (0x25CA) 项目符号 4级
     */
    public static final String LEVEL_SYMBOL_4 = String.valueOf((char) 0x25CA);
    /**
     * 项目符号字体
     */
    public static final String FONT_SYMBOL = "Wingdings 2";
    private static final Map<Class<? extends RenderData>, Character> RENDER_DATA_SYMBOL_MAP = new LinkedHashMap<>(6);
    private static final String REPLACE_1 = new String(new char[]{0X6C, 0XA0});
    private static final Configure COMPILE_CONFIGURE = newConfigure(null);

    static {
        /* 图片 '@'
         * 文本 '\0'
         * 表格 '#'
         * 列表 '*'
         * word文档模板 '+'
         */
        RENDER_DATA_SYMBOL_MAP.put(PictureRenderData.class, com.deepoove.poi.config.GramerSymbol.IMAGE.getSymbol());
        RENDER_DATA_SYMBOL_MAP.put(MiniTableRenderData.class, com.deepoove.poi.config.GramerSymbol.TABLE.getSymbol());
        RENDER_DATA_SYMBOL_MAP.put(NumbericRenderData.class, com.deepoove.poi.config.GramerSymbol.NUMBERIC.getSymbol());
        RENDER_DATA_SYMBOL_MAP.put(DocxProjectFileRenderData.class, com.deepoove.poi.config.GramerSymbol.DOCX_TEMPLATE.getSymbol());

        LEVEL_SYMBOL_STYLE.setFontFamily(FONT_SYMBOL);
        LEVEL_SYMBOL_STYLE.setBold(false);
        LEVEL_SYMBOL_STYLE.setFontSize(5);
    }

    /**
     * 字符串拦截特殊符号 (所有文本)
     *
     * @param str 任何字符串
     * @return 没有特殊符号的字符串
     */
    public static String stringFilter(String str) {
        if (str == null || str.isEmpty()) {
            return "";
        }
//        String replace = StringEscapeUtils.unescapeHtml(str)
        String replace = str
                .replace("</br>", "\n")
                .replace("<br>", "\n")
                .replace("<p>", "")
                .replace("</p>", "\n")
                .replace(REPLACE_1, "l ")
//                .replace("l","l")
                ;
//                .replace("&amp;nbsp;", " ")
//                .replace("&amp;", "&")
//                .replace("&nbsp;", " ")
//                .replace("nbsp;", " ")
//                .replace("&lt;", "<")
//                .replace("&gt;", ">");
        return replace;
    }

    /**
     * 前后回车
     *
     * @param str
     * @return
     */
    public static String wrapLine(String str) {
        if (str == null || str.isEmpty()) {
            return str;
        }
        str = str.trim();
        if (str.startsWith("<p>") && str.endsWith("</p>")) {
            return str;
        } else {
            return "<p>" + str + "</p>";
        }
    }

    public static Configure newConfigure(Context context) {
        Configure configure = Configure.newBuilder()
                //自定义字段渲染 docx如数据的数组为空,则写入空文本或删除文档元素的渲染
                .addPlugin(com.deepoove.poi.config.GramerSymbol.DOCX_TEMPLATE.getSymbol(), new CustomFieldRenderPolicy())
                //动态文本渲染, 用于支持动态元素类型
                .addPlugin(com.deepoove.poi.config.GramerSymbol.TEXT.getSymbol(), new DynamicTextRenderPolicy())
                .build();

        if (context != null) {
            List<RenderFinishRenderPolicy> renderFinishRenderPolicyList = context.getFinishRenderPolicyList();
            for (RenderFinishRenderPolicy each : renderFinishRenderPolicyList) {
                configure.customPolicy(each.tagName(), each);
            }
        }
        return configure;
    }

    public static void extendStyle(Style source, Style target) {
        if (source == null || target == null) {
            return;
        }
        if (target.getColor() == null) {
            target.setColor(source.getColor());
        }
        if (target.getFontFamily() == null) {
            target.setFontFamily(source.getFontFamily());
        }
        if (target.getFontSize() <= 0) {
            target.setFontSize(source.getFontSize());
        }
        if (target.isBold() == null) {
            target.setBold(source.isBold());
        }
        if (target.isItalic() == null) {
            target.setItalic(source.isItalic());
        }
        if (target.isStrike() == null) {
            target.setStrike(source.isStrike());
        }
        if (target.isUnderLine() == null) {
            target.setUnderLine(source.isUnderLine());
        }
    }

    public static boolean isSymbolItem(String text) {
        if (text != null) {
            for (int i = 0; i < text.length(); i++) {
                char c = text.charAt(i);
                if (c == ' ' || c == '-' || c == '_') {
                    continue;
                }
                if (c > ' ' && c <= 0XA0) {
                    return false;
                }
                if (c >= 0x4E00 && c <= 0x9FA5) {
                    return false;
                }
                return true;
            }
        }
        return false;
    }

    public static boolean isBlank(CharSequence str) {
        int strLen;
        if (str == null || (strLen = str.length()) == 0) {
            return true;
        }
        for (int i = 0; i < strLen; i++) {
            char c = str.charAt(i);
            if (c == '\n' || c == '\t') {
                return false;
            }
            if ((!Character.isWhitespace(c))) {
                return false;
            }
        }
        return true;
    }

    public static boolean isNotEmpty(Object object) {
        if (object == null) {
            return false;
        }
        if (object instanceof CharSequence) {
            return !isBlank((CharSequence) object);
        }
        if (object instanceof DocxProjectFileRenderData) {
            return true;
        }
        if (object instanceof CustomField) {
            return !((CustomField) object).isEmptyCustomField();
        }
        if (object.getClass().isArray()) {
            return Array.getLength(object) > 0;
        }
        if (object instanceof Collection) {
            return !((Collection) object).isEmpty();
        }
        if (object instanceof DocxRenderData) {
            return !((DocxRenderData) object).getDataModels().isEmpty();
        }
        if (object instanceof NumbericRenderData) {
            return !((NumbericRenderData) object).getNumbers().isEmpty();
        }
        if (object instanceof TextRenderData) {
            return isNotEmpty(((TextRenderData) object).getText());
        }
        if (object instanceof Map) {
            return !((Map) object).isEmpty();
        }
        return true;
    }

    public static boolean isAnyEmpty(Object... object) {
        for (Object str : object) {
            if (isEmpty(str)) {
                return true;
            }
        }
        return false;
    }

    public static boolean isEmpty(Object object) {
        return !isNotEmpty(object);
    }

    public static <T> void isNotEmpty(T object, Consumer<T> consumer) {
        if (isNotEmpty(object)) {
            consumer.accept(object);
        }
    }

    public static String toUpperCase(String input) {
        return StringUtil.toUpperCase(input);
    }

    /**
     * @param docMerges 待合并的文档
     * @param run       合并的位置
     * @param document
     * @return 合并后的文档
     * @throws Exception
     * @since 1.3.0
     */
    public static com.deepoove.poi.xwpf.NiceXWPFDocument merge(com.deepoove.poi.xwpf.NiceXWPFDocument document, List<com.deepoove.poi.xwpf.NiceXWPFDocument> docMerges, XWPFRun run) throws Exception {
        Context context = CONTEXT_LOCAL.get();
        XWPFParagraph paragraph = context.getMergePos();
        if (paragraph == null) {
            XmlCursor cursor;
            try {
                cursor = ((XWPFParagraph) run.getParent()).getCTP().newCursor();
                paragraph = document.insertNewParagraph(cursor);
            } catch (Exception e) {
                throw e;
            }
        }
        context.setMergePos(null);
        if (paragraph == null) {
            paragraph = document.getParagraphArray(0);
        }
        return merge(document, docMerges, paragraph);
    }

    public static IBodyElement getBodyElement(IBody body, int index) {
        if (index < 0) {
            return null;
        }
        List<IBodyElement> bodyElements = body.getBodyElements();
        IBodyElement bodyElement;
        if (index < bodyElements.size()) {
            bodyElement = bodyElements.get(index);
        } else {
            bodyElement = null;
        }
        return bodyElement;
    }

    /**
     * @param docMerges 待合并的文档
     * @param paragraph 合并的位置
     * @param document
     * @return 合并后的文档
     * @throws Exception
     * @since 1.3.0
     */
    public static NiceXWPFDocument merge(NiceXWPFDocument document, List<NiceXWPFDocument> docMerges, XWPFParagraph paragraph) throws Exception {
        return document.merge(docMerges.iterator(), paragraph);
    }

    private static void setText(TextRenderData data, XWPFRun run) {
        String text = data.getText();
        Style defaultStyle = data.getStyle();
        if (null == text) {
            clearText(run);
        } else if (text.isEmpty() || (text.length() == 1 && text.charAt(0) == '\n')) {
            //加行
            clearText(run);
            run.addBreak();
        } else {
            AST ast = JsoupAST.parsePart(text);
            List<? extends AST> nodes = ast.getChildren();
            if (nodes.isEmpty()) {
                nodes = ast.collect();
            }
            if (nodes.size() == 1) {
                nodes = Collections.singletonList(ast);
            }
            if (nodes.isEmpty()) {
                run.setText(text, 0);
            } else {
//                log.info("setText = {}, nodes={}",text,nodes);
                DocElement docElement = new DocElement(nodes.get(0), nodes, 0, data);
                setTextAndStyle(defaultStyle, run, docElement);
                for (int i = 1; i < nodes.size(); i++) {
                    AST node = nodes.get(i);
                    docElement.setNodeIndex(i);
                    docElement.setNode(node);
                    XWPFParagraph parent = (XWPFParagraph) run.getParent();
                    setTextAndStyle(defaultStyle, parent.createRun(), docElement);
                }
            }
        }
    }

    private static void setTextAndStyle(Style defaultStyle, XWPFRun rootRun, DocElement docElement) {
        TextRenderData data = docElement.getData();
        if (data instanceof TextRenderFilterData && ((TextRenderFilterData) data).getAstDocConsumer() != null) {
            docElement.clear();
            docElement.add(rootRun, "", defaultStyle);
            ((TextRenderFilterData) data).getAstDocConsumer().accept(docElement);
            for (DocElement.RunWrapper wrapper : docElement) {
                XWPFRun run = wrapper.getRun();
                StyleUtils.styleRun(run, wrapper.getStyle());
                String content = wrapper.getContent();
                setText(run, content);
            }
        } else {
            setText(rootRun, docElement.getNode().getContent());
        }
    }

    private static void setText(XWPFRun run, String content) {
        if (content == null || content.isEmpty()) {
            clearText(run);
        } else {
            String[] split = content.split("\\n");
            run.setText(split[0], 0);
            for (int i = 1; i < split.length; i++) {
                run.addBreak();//\n的换行
                run.setText(split[i], i);
            }
        }
    }

    public static NiceXWPFDocument mergeTable(List<NiceXWPFDocument> docs) {
        if (docs.size() > 1) {
            NiceXWPFDocument firstDocument = docs.get(0);
            XWPFTable firstTable = firstDocument.getTables().get(firstDocument.getTables().size() - 1);
            for (NiceXWPFDocument doc : new ArrayList<>(docs)) {
                if (doc == firstDocument) {
                    continue;
                }
                for (XWPFTable table : doc.getTables()) {
                    for (XWPFTableRow row : table.getRows()) {
                        firstTable.addRow(row);
                    }
                }
                docs.remove(doc);
            }
        }
        return null;
    }

    private static void onRenderBefore(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        Stack<Context.Frame> stack = CONTEXT_LOCAL.get().getRenderDataStack();
        Context.Frame frame = new Context.Frame(eleTemplate, data, template);
        if(data instanceof RenderStackFrameListener){
            Consumer3 listener = ((RenderStackFrameListener) data).getRenderBeforeListener();
            if(listener != null) {
                Class<?> dataType = findGenericInterfacesType(data, RenderStackFrameListener.class, "DATA");
                if (dataType.isAssignableFrom(data.getClass())) {
                    listener.accept(data, frame, stack);
                }
            }
        }
        stack.push(frame);
    }

    private static void onRenderAfter(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
        Stack<Context.Frame> stack = CONTEXT_LOCAL.get().getRenderDataStack();
        Context.Frame frame = stack.pop();
        if(data instanceof RenderStackFrameListener){
            Consumer3 listener = ((RenderStackFrameListener) data).getRenderAfterListener();
            if(listener != null) {
                Class<?> dataType = findGenericInterfacesType(data, RenderStackFrameListener.class, "DATA");
                if (dataType.isAssignableFrom(data.getClass())) {
                    listener.accept(data, frame, stack);
                }
            }
        }
    }

    public static Class<?> findGenericInterfacesType(Object object, Class<?> parametrizedSuperclass, String typeParamName) {
        Class<?> thisClass = object.getClass();
        Class currentClass = thisClass;
        while (true) {
            for (Type genericInterface : currentClass.getGenericInterfaces()) {
                if (genericInterface instanceof ParameterizedType) {
                    ParameterizedType parameterizedType = (ParameterizedType) genericInterface;

                    Type rawType = parameterizedType.getRawType();
                    if (!(rawType instanceof GenericDeclaration)) {
                        continue;
                    }
                    if (rawType != parametrizedSuperclass) {
                        continue;
                    }
                    int typeParamIndex = -1;
                    TypeVariable<?>[] typeParams = ((GenericDeclaration) rawType).getTypeParameters();
                    for (int i = 0; i < typeParams.length; ++i) {
                        if (typeParamName.equals(typeParams[i].getName())) {
                            typeParamIndex = i;
                            break;
                        }
                    }
                    Type actualTypeArgument = parameterizedType.getActualTypeArguments()[typeParamIndex];
                    if (actualTypeArgument instanceof Class) {
                        return (Class<?>) actualTypeArgument;
                    }
                }
            }
            currentClass = currentClass.getSuperclass();
            if (currentClass == null || currentClass == Object.class) {
                throw new IllegalStateException("cannot determine the type of the type parameter '" + typeParamName + "': " + thisClass);
            }
        }
    }

    private static String buildSource(Configure configure, Character symbol, String tagName) {
        if (symbol == null) {
            return configure.getGramerPrefix() + tagName + configure.getGramerSuffix();
        }
        return configure.getGramerPrefix() + symbol + tagName + configure.getGramerSuffix();
    }

    private static List<NiceXWPFDocument> getMergedDocxs(ElementTemplate eleTemplate, DocxRenderData data, Configure configure, MergedOverrideMethod overrideMethod) {
        List<NiceXWPFDocument> docs = new ArrayList<>();
        byte[] docx = data.getDocx();
//        Supplier<byte[]> supplier = CONTEXT_LOCAL.get().getTemplate().getByLength(docx.length);
        List<?> dataList = data.getDataModels();
        if (null != dataList) {
            for (Object o : dataList) {
                byte[] eachDocx = null;
                if (o instanceof Map) {
                    Function<Map, byte[]> choseRowDocx = (Function<Map, byte[]>) ((Map) o).get(CustomField.FIELD_CHOSEROWDOCX);
                    if (choseRowDocx != null) {
                        eachDocx = choseRowDocx.apply((Map) o);
                    }
                }
                if (eachDocx == null) {
                    eachDocx = docx;
                }
                XWPFTemplate template = DeepooveUtil.compile(eachDocx, configure, o);
                Callable<NiceXWPFDocument> rawMethod = () -> DeepooveUtil.render(template, o).getXWPFDocument();
                NiceXWPFDocument document;
                try {
                    if (overrideMethod != null) {
                        onRenderBefore(eleTemplate, o, template);
                        try {
                            document = overrideMethod.apply(template, o, docs, rawMethod);
                        } finally {
                            onRenderAfter(eleTemplate, o, template);
                        }
                    } else {
                        document = rawMethod.call();
                    }
                } catch (Exception e) {
                    document = template.getXWPFDocument();
                    log.error("getMergedDocxs error. data={}, error={}", o, e.toString(), e);
                }
                removeLastLineIfNeed(template.getXWPFDocument(), o);
                docs.add(document);
            }
        }
        return docs;
    }

    public static void removeLastLineIfNeed(NiceXWPFDocument document, Object renderData) {
        removeLastLineIfNeed(document, renderData, false);
    }

    public static void removeLastLineIfNeed(NiceXWPFDocument document, Object renderData, boolean removeEmptyLine) {
        //删掉最后一个多余自动追加的空行. 王子豪 2019年11月22日 15:20:09
        int bodySize = document.getBodyElements().size();
        if (bodySize != 0) {
            int lastBodyIndex = bodySize - 1;
            IBodyElement lastBody = document.getBodyElements().get(lastBodyIndex);
            if (isNeedRemove(lastBody)) {
                removeBodyElement(document, lastBodyIndex, "Merged after remove", renderData);
            } else if (removeEmptyLine) {
                removeEmptyLine(lastBody);
            }
        }
    }

    public static RenderPolicy getPolicy(Configure configure, String tagName, Character sign) {
        RenderPolicy policy = configure.getCustomPolicy(tagName);
        if (policy == null) {
            policy = configure.getDefaultPolicy(sign);
        }
        return policy;
    }

    public static int removeEmptyLine(IBodyElement bodyElement) {
        return removeEmptyLine(bodyElement, false);
    }

    public static void removeFirstEmptyLine(List<XWPFTable> tableList) {
        List<XWPFTable> removeEmptyLineTableList = new ArrayList<>();
        if (tableList != null) {
            for (XWPFTable table : tableList) {
                for (XWPFTableRow row : table.getRows()) {
                    List<XWPFTableCell> cells = row.getTableCells();
                    if (cells.isEmpty()) {
                        continue;
                    }
                    boolean isNotBlank = cells.get(0).getBodyElements().stream().anyMatch(e -> !isBlankLine(e));
                    if (isNotBlank) {
                        break;
                    }
                    row.removeCell(0);
                }
            }
//            for (XWPFTable table : removeEmptyLineTableList) {
//                removeEmptyLine(table, true, 1);
//            }
        }
    }

    public static int removeEmptyLine(IBodyElement bodyElement, boolean related) {
        return removeEmptyLine(bodyElement, related, Integer.MAX_VALUE);
    }

    public static int removeEmptyLine(IBodyElement bodyElement, boolean related, int maxRemove) {
        int removeCount = 0;
        if (bodyElement instanceof XWPFTable) {
            XWPFTable table = (XWPFTable) bodyElement;
            String gramerPrefix = COMPILE_CONFIGURE.getGramerPrefix();
            List<XWPFTableRow> removeRows = new ArrayList<>();
            for (XWPFTableRow row : new ArrayList<>(table.getRows())) {
                for (XWPFTableCell cell : new ArrayList<>(row.getTableCells())) {
                    for (XWPFParagraph paragraph : new ArrayList<>(cell.getParagraphs())) {
                        if (paragraph.getRuns().isEmpty() ||
                                DeepooveUtil.bodyToString(paragraph).startsWith(gramerPrefix)) {
                            if (removeCount >= maxRemove) {
                                break;
                            }
                            if (DeepooveUtil.removeBodyElement(cell, paragraph, null, null)) {
                                removeCount++;
                            }
                        }
                    }
                    if (related && cell.getBodyElements().isEmpty()) {
                        row.getTableCells().remove(cell);
                    }
                    if (removeCount >= maxRemove) {
                        break;
                    }
                }
                if (related && row.getTableCells().isEmpty()) {
                    removeRows.add(row);
                }
                if (removeCount >= maxRemove) {
                    break;
                }
            }
            removeRows(table, removeRows);
        }
        return removeCount;
    }

    private static void removeRows(XWPFTable table, List<XWPFTableRow> removeList) {
        for (XWPFTableRow remove : removeList) {
            for (int i = 0; i < table.getRows().size(); i++) {
                if (table.getRow(i) == remove) {
                    table.removeRow(i);
                    break;
                }
            }
        }
    }

    public static String bodyToString(XWPFRun body) {
        try {
            return body.text();
        } catch (Exception e) {
            log.warn("bodyToString(XWPFRun) error {}",e.toString(),e);
            return "";
        }
    }

    public static String bodyToString(IBodyElement body) {
        if (body instanceof XWPFTable) {
            try {
                return ((XWPFTable) body).getText();
            }catch (Exception e){
                log.warn("bodyToString(XWPFTable) error {}",e.toString(),e);
                return "";
            }
        }
        if (body instanceof XWPFParagraph) {
            try {
                return ((XWPFParagraph) body).getText();
            } catch (Exception e) {
                log.warn("bodyToString(XWPFParagraph) error {}",e.toString(),e);
                return "";
            }
        }
        try {
            return body.toString();
        }catch (Exception e){
            log.warn("bodyToString(IBodyElement) error {}",e.toString(),e);
            return "";
        }
    }

    /**
     * 根据值找内容
     *
     * @param body
     * @param value 值
     * @param type  值类型
     * @param asc   true=正序找,false=倒序找
     * @return 内容的下标
     */
    public static int findBodyIndexByValue(IBody body, String value, Class type, boolean asc) {
        List<IBodyElement> bodyElements = body.getBodyElements();
        if (asc) {
            for (int i = bodyElements.size() - 1; i >= 0; i--) {
                IBodyElement bodyElement = bodyElements.get(i);
                if (type == bodyElement.getClass()) {
                    String s = bodyToString(bodyElement);
                    if (Objects.equals(s, value)) {
                        return i;
                    }
                }
            }
        } else {
            for (int i = 0; i < bodyElements.size(); i++) {
                IBodyElement bodyElement = bodyElements.get(i);
                if (type == bodyElement.getClass()) {
                    String s = bodyToString(bodyElement);
                    if (Objects.equals(s, value)) {
                        return i;
                    }
                }
            }
        }
        return -1;
    }

    public static <T> List<T> findBodyIndexByStartWithValue(IBody body, String startWithValue, Class<T> type, boolean asc) {
        List<T> list = new ArrayList<>();
        List<IBodyElement> bodyElements = body.getBodyElements();
        if (asc) {
            for (int i = bodyElements.size() - 1; i >= 0; i--) {
                IBodyElement bodyElement = bodyElements.get(i);
                if (type == bodyElement.getClass()) {
                    String s = bodyToString(bodyElement);
                    if (s.startsWith(startWithValue)) {
                        list.add((T) bodyElement);
                    }
                }
            }
        } else {
            for (int i = 0; i < bodyElements.size(); i++) {
                IBodyElement bodyElement = bodyElements.get(i);
                if (type == bodyElement.getClass()) {
                    String s = bodyToString(bodyElement);
                    if (s.startsWith(startWithValue)) {
                        list.add((T) bodyElement);
                    }
                }
            }
        }
        return list;
    }

    public static int findBodyIndexByBody(IBody document, Object findBody) {
        int index = 0;
        for (IBodyElement body : document.getBodyElements()) {
            if (findBody == body) {
                return index;
            }
            index++;
        }
        return -1;
    }

    public static int findBodyIndexByBody(XWPFParagraph document, Object findBody) {
        int index = 0;
        for (IRunElement body : document.getRuns()) {
            if (findBody == body) {
                return index;
            }
            index++;
        }
        return -1;
    }

    private static int findBodyIndexByTagName(IBody document, String tagSource) {
        int index = 0;
        for (IBodyElement body : document.getBodyElements()) {
            if (isEqualsTag(body, tagSource)) {
                return index;
            }
            index++;
        }
        return -1;
    }

    public static List<XWPFRun> findRunList(List<XWPFTable> tables, Predicate<String> test) {
        List<XWPFRun> runList = new ArrayList<>();
        for (XWPFTable table : tables) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (IBodyElement bodyElement : cell.getBodyElements()) {
                        if (!(bodyElement instanceof XWPFParagraph)) {
                            continue;
                        }
                        for (XWPFRun run : ((XWPFParagraph) bodyElement).getRuns()) {
                            String text = bodyToString(run);
                            if (text.isEmpty()) {
                                continue;
                            }
                            if (test.test(text)) {
                                runList.add(run);
                            }
                        }
                    }
                }
            }
        }
        return runList;
    }

    private static String stringReplaceXML(String str) {
        if (str == null || str.isEmpty()) {
            return str;
        }
        return str.replace("\r", "")
                .replace("\t", "");
    }

    private static boolean isEqualsTag(IBodyElement body, String tagName) {
        if (body instanceof XWPFTable) {
            String text = bodyToString(body);
            if (text == null) {
                return tagName == null;
            }
            String textReplace = stringReplaceXML(text);
            if (textReplace.equals(tagName)) {
                return true;
            }
        }
        if (body instanceof XWPFParagraph) {
            for (XWPFRun run : ((XWPFParagraph) body).getRuns()) {
                String text = bodyToString(run);
                boolean isTag = text != null && text.contains(tagName);
                if (isTag) {
                    return true;
                }
            }
        }
        return false;
    }

    private static void clearText(XWPFRun run) {
        try {
            run.setText("", 0);
        } catch (Exception e) {
            //skip
        }
    }

    public static boolean removeBodyElement(XWPFDocument document, Object remove, String tagName, Object renderData) {
        int bodyIndexByBody = findBodyIndexByBody(document, remove);
        return removeBodyElement(document, bodyIndexByBody, tagName, renderData);
    }

    public static boolean removeBodyElement(XWPFParagraph document, Object remove, String tagName, Object renderData) {
        int bodyIndexByBody = findBodyIndexByBody(document, remove);
        return document.removeRun(bodyIndexByBody);
    }

    public static boolean removeBodyElement(XWPFDocument document, IRunBody remove, String tagName, Object renderData) {
        int bodyIndexByBody = findBodyIndexByBody(document, remove);
        return removeBodyElement(document, bodyIndexByBody, tagName, renderData);
    }

    public static boolean removeBodyElement(XWPFDocument document, int index, String tagName, Object renderData) {
        List<IBodyElement> bodyElements = document.getBodyElements();
        if (index >= 0 && index < bodyElements.size()) {
            IBodyElement removeBody = bodyElements.get(index);
            String bodyToString;
            try {
                bodyToString = bodyToString(removeBody);
            } catch (Exception e) {
                bodyToString = "disappeared";
            }
            log.trace("removeBodyElement = {}, tagName='{}', data={}", bodyToString, tagName, renderData);
            try {
                document.removeBodyElement(index);
                return true;
            } catch (IndexOutOfBoundsException e) {
                //解决poi 删除失败的 bug
                Iterator<IBodyElement> iterator = document.getBodyElementsIterator();
                while (iterator.hasNext()) {
                    if (removeBody == iterator.next()) {
                        iterator.remove();
                    }
                }
                return true;
            }
        }
        return false;
    }

    public static boolean removeBodyElement(XWPFTableCell document, IRunBody remove, String tagName, Object renderData) {
        int bodyIndexByBody = findBodyIndexByBody(document, remove);
        return removeBodyElement(document, bodyIndexByBody, tagName, renderData);
    }

    public static boolean removeBodyElement(XWPFTableCell document, int index, String tagName, Object renderData) {
        List<IBodyElement> bodyElements = document.getBodyElements();
        if (index >= 0 && index < bodyElements.size()) {
            IBodyElement removeBody = bodyElements.get(index);
            String bodyToString;
            try {
                bodyToString = bodyToString(removeBody);
            } catch (Exception e) {
                bodyToString = "disappeared";
            }
            log.trace("removeBodyElement = {}, tagName='{}', data={}", bodyToString, tagName, renderData);
            try {
                document.removeParagraph(index);
                try {
                    List<IBodyElement> bodyElementList = (List<IBodyElement>) BeanUtil.getFieldValue("bodyElements", document);
                    if(bodyElementList == null || bodyElementList.isEmpty() || bodyElementList.getClass().getName().contains("Unmodifiable")){
                        return false;
                    }
                    bodyElementList.remove(removeBody);
                } catch (IllegalAccessException | NoSuchFieldException e) {
                    e.printStackTrace();
                }
                return true;
            } catch (IndexOutOfBoundsException e) {
                throw e;
            }
        }
        return false;
    }

    /**
     * 是否需要删除,把空行删掉
     *
     * @param body body
     * @return true=需要删,false=不删
     */
    private static boolean isNeedRemove(IBodyElement body) {
        return isBlankLine(body);
    }

    /**
     * 是否是空行
     *
     * @param body body
     * @return true=是,false=不是
     */
    public static boolean isEmptyLine(IBodyElement body) {
        if (body instanceof XWPFTable) {
            String text = bodyToString(body);
            String textReplace = stringReplaceXML(text);
            return textReplace == null || textReplace.isEmpty();
        }
        if (body instanceof XWPFParagraph) {
            for (XWPFRun run : ((XWPFParagraph) body).getRuns()) {
                String text = stringFilter(bodyToString(run));
                if (text != null && !text.isEmpty()) {
                    return false;
                }
            }
            return true;
        }
        return false;
    }

    /**
     * 是否是空行
     *
     * @param body body
     * @return true=是,false=不是
     */
    public static boolean isBlankLine(IBodyElement body) {
        if (body instanceof XWPFTable) {
            String text = bodyToString(body);
            String textReplace = stringReplaceXML(text);
            return StringUtil.isBlank(textReplace);
        }
        if (body instanceof XWPFParagraph) {
            for (XWPFRun run : ((XWPFParagraph) body).getRuns()) {
                String text = stringFilter(bodyToString(run));
                if (StringUtil.isNotBlank(text)) {
                    return false;
                }
            }
            return true;
        }
        return false;
    }

    /**
     * 是否是空行
     *
     * @param cell cell
     * @return true=是,false=不是
     */
    public static boolean isEmptyCell(XWPFTableCell cell) {
        for (IBodyElement bodyElement : cell.getBodyElements()) {
            if (!isEmptyLine(bodyElement)) {
                return false;
            }
        }
        return true;
    }

    /**
     * 加回车行(如果不满足回车数量的话)
     *
     * @param document        文档
     * @param beginIndex      这个下标前的回车数量开始计算
     * @param fixedEmptyCount 要求的回车数量
     * @return 加了多少回车
     */
    public static int addLineIfNeed(IBody document, int beginIndex, int fixedEmptyCount) {
        IBodyElement eduExpsTitleBodyElement = getBodyElement(document, beginIndex);
        IBodyElement eduExpsTitleBodyElement1 = getBodyElement(document, beginIndex - 1);
        if (eduExpsTitleBodyElement == null || eduExpsTitleBodyElement1 == null) {
            return 0;
        }
        int emptyCount = 0;
        for (int i = 1; i <= fixedEmptyCount; i++) {
            IBodyElement iBodyElement = getBodyElement(document, beginIndex - i);
            if (iBodyElement == null) {
                continue;
            }
            if (isBlankLine(iBodyElement)) {
                emptyCount++;
            }
        }
        int addEmptyLineCount = fixedEmptyCount - emptyCount;
        for (int i = 0; i < addEmptyLineCount; i++) {
            XmlCursor cursor = ((XWPFParagraph) eduExpsTitleBodyElement1).getCTP().newCursor();
            XWPFParagraph paragraph = document.insertNewParagraph(cursor);
            paragraph.createRun();
        }
        return addEmptyLineCount;
    }

    /**
     * 新添加一行回车 (在每个文档中间加回车)
     *
     * @param documents     文档
     * @param appendEndLine 是否在最后追加一行回车
     * @return 新的一行
     */
    public static void addLine(List<NiceXWPFDocument> documents, boolean appendEndLine) {
        if (documents.size() > 0) {
            for (NiceXWPFDocument document : documents) {
                addLine(document);
            }
            if (appendEndLine) {
                NiceXWPFDocument lastDocument = documents.get(documents.size() - 1);
                addLine(lastDocument);
            }
        }
    }

    /**
     * 新添加一行回车 (加在最后)
     *
     * @param document 文档
     * @return 新的一行
     */
    public static XWPFRun addLine(XWPFDocument document) {
        XWPFRun run = document.createParagraph().createRun();
        run.setText("");
        return run;
    }

    /**
     * 清空全部空行
     *
     * @param document        文档
     * @param debugInfo       调试信息
     */
    public static void clearEmptyLine(NiceXWPFDocument document, String debugInfo) {
        removeEmptyFixedLineIfNeed(document, Integer.MAX_VALUE, 1, Integer.MAX_VALUE, debugInfo);
    }

    /**
     * 固定回车行数 (如果不是大于回车数量, 则删除一行)
     *
     * @param document        文档
     * @param index           从第几个开始递减扫描
     * @param fixedEmptyCount 固定回车行数
     * @param maxRemoveCount  最多删除多少个回车
     * @param debugInfo       调试信息
     */
    public static void removeEmptyFixedLineIfNeed(NiceXWPFDocument document, int index, int fixedEmptyCount, int maxRemoveCount, String debugInfo) {
        List<IBodyElement> bodyElements = document.getBodyElements();
        int size = bodyElements.size();
        int emptyCount = 0;
        for (int i = 0; i < size; i++) {
            if (isBlankLine(bodyElements.get(i))) {
                emptyCount++;
            }
        }
        if (emptyCount <= fixedEmptyCount) {
            return;
        }
        int removeCount = emptyCount - fixedEmptyCount;
        boolean lastEmptyFlag = false;
        List<IBodyElement> removeList = new ArrayList<>();
        for (int i = Math.min(size - 1, index); i > 0; i--) {
            IBodyElement bodyElement = bodyElements.get(i);
            if (lastEmptyFlag && isBlankLine(bodyElement)) {
                removeList.add(bodyElement);
                if (removeList.size() >= maxRemoveCount) {
                    break;
                } else {
                    continue;
                }
            }
            for (int j = 0; j < fixedEmptyCount; j++) {
                int findIndex = i - j;
                if (findIndex < 0) {
                    break;
                }
                IBodyElement findBodyElement = bodyElements.get(findIndex);
                boolean cEmpty = isBlankLine(findBodyElement);
                i = findIndex;
                if (cEmpty) {
                    lastEmptyFlag = true;
                } else {
                    lastEmptyFlag = false;
                    break;
                }
            }
        }
        for (IBodyElement bodyElement : removeList) {
            removeBodyElement(document, bodyElement, null, debugInfo);
        }
    }

    public static XWPFTable findFirstTableByDocument(List<NiceXWPFDocument> documents) {
        XWPFTable table;
        for (NiceXWPFDocument document : documents) {
            if ((table = findFirstTable(document.getTables())) != null) {
                return table;
            }
        }
        return null;
    }

    public static XWPFTable findFirstTable(List<XWPFTable> tables) {
        return tables == null || tables.isEmpty() ? null : tables.get(0);
    }

    public static void mergeTable(XWPFTable table, List<XWPFTable> tables, XWPFDocument document, List<IBodyElement> removeList) {
        if (tables.isEmpty()) {
            return;
        }
        for (XWPFTable eachTable : tables) {
            List<XWPFTableRow> rows = new ArrayList<>(eachTable.getRows());
            for (int i = 0; i < rows.size(); i++) {
                table.addRow(rows.get(i));
            }
        }
        if (removeList != null) {
            for (IBodyElement bodyElement : new ArrayList<>(removeList)) {
                if (table == bodyElement) {
                    continue;
                }
                removeBodyElement(document, bodyElement, null, null);
            }
        }
    }

    public static void eachRuns(IBody iBody, Consumer<XWPFParagraph> consumer) {
        XWPFRun lastRun = null;
        for (IBodyElement bodyElement : iBody.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph) {
                XWPFParagraph paragraph = ((XWPFParagraph) bodyElement);
                consumer.accept(paragraph);
            } else if (bodyElement instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) bodyElement).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        eachRuns(cell, consumer);
                    }
                }
            }
        }
    }

    public static XWPFRun getFirstRun(IBody iBody) {
        for (IBodyElement bodyElement : iBody.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph) {
                List<XWPFRun> runs = ((XWPFParagraph) bodyElement).getRuns();
                if (isEmpty(runs)) {
                    continue;
                }
                return runs.get(runs.size() - 1);
            } else if (bodyElement instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) bodyElement).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        XWPFRun cellLastRun = getFirstRun(cell);
                        if (cellLastRun != null) {
                            return cellLastRun;
                        }
                    }
                }
            }
        }
        return null;
    }

    public static XWPFRun getLastRun(IBody iBody) {
        XWPFRun lastRun = null;
        for (IBodyElement bodyElement : iBody.getBodyElements()) {
            if (bodyElement instanceof XWPFParagraph) {
                List<XWPFRun> runs = ((XWPFParagraph) bodyElement).getRuns();
                if (isEmpty(runs)) {
                    continue;
                }
                lastRun = runs.get(runs.size() - 1);
            } else if (bodyElement instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) bodyElement).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        XWPFRun cellLastRun = getLastRun(cell);
                        if (cellLastRun != null) {
                            lastRun = cellLastRun;
                        }
                    }
                }
            }
        }
        return lastRun;
    }

    public static RunTemplate getRunTemplate(List<ElementTemplate> templates, String tagName) {
        for (ElementTemplate template : templates) {
            if (Objects.equals(template.getTagName(), tagName)) {
                return (RunTemplate) template;
            }
        }
        throw new IllegalArgumentException("tagName [" + tagName + "] not exist");
    }

    /**
     * 渲染的入口方法
     *
     * @param template
     * @param bean
     * @param context
     */
    public static void renderMain(XWPFTemplate template, Object bean, Context context) {
        try {
            render(template, bean);
            if (context != null) {
                //二次渲染, 删除多余的行. 那些需要等全部完成后处理逻辑的模版
                List<RenderFinishRenderPolicy> finishRenderPolicyList = context.getFinishRenderPolicyList();
                finishRenderPolicyList.forEach(RenderFinishRenderPolicy::onRenderFinish);
                render(template, bean);
            }
        } finally {
            if (context != null) {
                log.info("renderMain end. poolHit = {}, compileTotal={}", context.getCompileHit(), context.getCompileTotal());
            }
        }
    }

    private static XWPFTemplate render(XWPFTemplate template, Object bean) {
        onRenderBefore(null, bean, template);
        try {
            Map datas = BeanMap.toMap(bean);
            if (null == template) {
                throw new POIXMLException("Template is null, should be setted first.");
            }
            List<ElementTemplate> elementTemplates = (List) template.getElementTemplates();
            if (null == elementTemplates || elementTemplates.isEmpty() || datas.isEmpty()) {
                return template;
            }
            Configure config = template.getConfig();

            int docxNum = 0;
            for (ElementTemplate runTemplate : elementTemplates) {
                putTagRun(datas, runTemplate.getTagName(), runTemplate);
            }
            for (ElementTemplate runTemplate : elementTemplates) {
                String tagName = runTemplate.getTagName();
                RenderPolicy policy = getPolicy(config, tagName, runTemplate.getSign());
                if (null == policy) {
                    throw new RenderException("cannot find render policy: [" + tagName + "]");
                }

                boolean isDocx;
                Object renderData = datas.get(tagName);
                if (renderData != null && policy instanceof DynamicTextRenderPolicy) {
                    isDocx = ((DynamicTextRenderPolicy) policy).isDocx(renderData.getClass(), config, tagName);
                } else {
                    isDocx = policy instanceof DocxRenderPolicy;
                }

                if (isDocx) {
                    docxNum++;
                    continue;
                }
                policy.render(runTemplate, renderData, template);
            }
            try {
                if (docxNum >= 1) {
                    template.reload(template.getXWPFDocument().generate());
                }
                for (int i = 0; i < docxNum; i++) {
                    renderDocx(template, datas);
                }
            } catch (Exception e) {
                log.error("render docx error={}", e.toString(), e);
            }
            return template;
        } finally {
            onRenderAfter(null, bean, template);
        }
    }

    private static void renderDocx(XWPFTemplate template, Map datas) {
        List<ElementTemplate> elementTemplates = (List) template.getElementTemplates();
        if (null == elementTemplates || elementTemplates.isEmpty() || datas.isEmpty()) {
            return;
        }
        Configure config = template.getConfig();
        for (ElementTemplate runTemplate : elementTemplates) {
            String tagName = runTemplate.getTagName();
            boolean isDocx;
            RenderPolicy policy = getPolicy(config, tagName, runTemplate.getSign());
            Object renderData = datas.get(tagName);
            if (renderData != null && policy instanceof DynamicTextRenderPolicy) {
                isDocx = ((DynamicTextRenderPolicy) policy).isDocx(renderData.getClass(), config, tagName);
            } else {
                isDocx = policy instanceof DocxRenderPolicy;
            }
            if (isDocx) {
                putTagRun(datas, tagName, runTemplate);
                policy.render(runTemplate, renderData, template);

                boolean isEmpty;
                if (renderData instanceof DocxRenderData) {
                    isEmpty = ((DocxRenderData) renderData).getDataModels().isEmpty();
                } else {
                    isEmpty = isEmpty(renderData);
                }
                if (isEmpty) {
                    elementTemplates.remove(runTemplate);
                    renderDocx(template, datas);
                }
                return;
            }
        }
    }

    private static void putTagRun(Map map, String tagName, ElementTemplate runTemplate) {
        CONTEXT_LOCAL.get().putTagRun(tagName, runTemplate);
        try {
            map.put(getTagRunKey(tagName), runTemplate);
        } catch (UnsupportedOperationException ig) {
            //skip
        }
    }

    public static String getTagRunKey(String tagName) {
        return tagName + "Run";
    }

    public static XWPFTemplate compile(byte[] file, Configure config, Object data) {
        XWPFTemplate template = directCompile(file, config);
        if (data instanceof AutowiredTemplate) {
            ((AutowiredTemplate) data).setTemplate(template);
        }
        return template;
    }

    public static XWPFTemplate directCompile(byte[] file, Configure config) {
        return XWPFTemplate.compile(new ByteArrayInputStream(file), config);
    }

    public enum CustomStyle {
        /**/
        SINGLE_TEXT, TABLE_ROW, TABLE_TABLE_ROW
    }

    public interface AutowiredTemplate {
        void setTemplate(XWPFTemplate template);
    }
    public interface RenderStackFrameListener<DATA> {
        Consumer3<DATA, Context.Frame, Stack<Context.Frame>> getRenderBeforeListener();
        Consumer3<DATA, Context.Frame, Stack<Context.Frame>> getRenderAfterListener();
    }
    @FunctionalInterface
    public interface Consumer3<T1,T2,T3> {
        void accept(T1 t1, T2 t2, T3 t3);
    }

    /**
     * 处理文档合并的逻辑
     */
    @FunctionalInterface
    public interface MergedOverrideMethod {
        NiceXWPFDocument apply(XWPFTemplate template, Object data, List<NiceXWPFDocument> docs, Callable<NiceXWPFDocument> rawMergedMethod) throws Exception;
    }

    /**
     * 文档元素（可以解释xml，html语言字符串，并动态映射到doc）
     * 它自身是多个元素的集合
     *
     * @see DocElement#node 解析字符串数据后的抽象语法树节点
     * @see RunWrapper#setContent(String) 为当前元素设置数据
     * @see RunWrapper#setStyle(Style) 为当前元素设置样式
     * @see DocElement#createRun() 创建新元素
     */
    @Getter
    @Setter
    @AllArgsConstructor
    public static class DocElement extends ArrayList<DocElement.RunWrapper> {
        private AST node;
        private List<? extends AST> nodes;
        private int nodeIndex;
        private TextRenderData data;

        /**
         * 填充表格一行的数据
         *
         * @param table
         * @param row     第几行
         * @param rowData 行数据：确保行数据的大小不超过表格该行的单元格数量
         */
        public static void renderRow(XWPFTable table, int row, List<String> rowData) {
            if (null == rowData || rowData.size() <= 0) {
                return;
            }
            int i = 0;
            XWPFTableCell cell = null;
            for (String cellData : rowData) {
                cell = table.getRow(row).getCell(i);
                String[] fragment = cellData.split("\\n");
                CTTc ctTc = cell.getCTTc();
                CTP ctP = (ctTc.sizeOfPArray() == 0) ? ctTc.addNewP() : ctTc.getPArray(0);
                XWPFParagraph par = new XWPFParagraph(ctP, cell);
                XWPFRun run = par.createRun();
                run.setText(fragment[0]);
                for (int j = 1; j < fragment.length; j++) {
                    XWPFParagraph addParagraph = cell.addParagraph();
                    run = addParagraph.createRun();
                    run.setText(fragment[j]);
                }
                i++;
            }
        }

        /**
         * 获取是项目符号的所有元素
         *
         * @return 项目符号
         */
        public List<XWPFRun> getSymbolItemList() {
            List<XWPFRun> runList = DeepooveUtil.findRunList(root().getTables(), DeepooveUtil::isSymbolItem);
            return runList;
        }

        /**
         * 设置第一个项目符号
         *
         * @param symbolItem 项目符号
         * @return 是否设置成功。 true=成功
         */
        public boolean setFirstSymbolItem(String symbolItem) {
            if (symbolItem != null) {
                List<XWPFRun> runList = getSymbolItemList();
                if (runList.size() > 0) {
                    XWPFRun symbolItemRun = runList.get(0);
                    setSymbolStyle(symbolItemRun);
                    symbolItemRun.setFontSize(10);
                    symbolItemRun.setText(symbolItem, 0);
                    return true;
                }
            }
            return false;
        }

        void setSymbolStyle(XWPFRun run) {
            StyleUtils.styleRun(run, LEVEL_SYMBOL_STYLE);
        }

        public RunWrapper createRun() {
            RunWrapper root = root();
            XWPFParagraph parent = (XWPFParagraph) root.run.getParent();
            return add(parent.createRun(), "", root.style);
        }

        public RunWrapper createCellRun() {
            RunWrapper root = root();
            XWPFParagraph parent = (XWPFParagraph) root.run.getParent();
            XWPFTableRow row = ((XWPFTableCell) (parent).getBody()).getTableRow();
            XWPFTableCell cell = row.addNewTableCell();
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            return add(paragraph.createRun(), "", root.style);
        }

        public RunWrapper add(XWPFRun run, String content, Style style) {
            RunWrapper wrapper = new RunWrapper(this, run, content, style);
            add(wrapper);
            return wrapper;
        }

        public RunWrapper root() {
            if (isEmpty()) {
                throw new IllegalStateException("RunWrapper size = 0, not root element");
            }
            return get(0);
        }

        public Style getStyle(int index) {
            return get(index).style;
        }

        public void setStyle(int index, Style style) {
            get(index).style = style;
        }

        public String getContent(int index) {
            return get(index).content;
        }

        public void setContent(int index, String content) {
            get(index).content = content;
        }

        @AllArgsConstructor
        @Getter
        @Setter
        public static class RunWrapper {
            private DocElement parent;
            private XWPFRun run;
            private String content;
            private Style style;

            public RunWrapper prev() {
                int i = parent.indexOf(this);
                if (i >= 1 && i < parent.size()) {
                    return parent.get(i - 1);
                }
                return null;
            }

            public List<XWPFTable> getTables() {
                return run.getDocument().getTables();
            }

            public RunWrapper next() {
                int i = parent.indexOf(this);
                if (i >= 0 && i < parent.size()) {
                    return parent.get(i + 1);
                }
                return null;
            }

            public void symbol() {
                parent.setSymbolStyle(run);
            }

            @Override
            public String toString() {
                return content;
            }
        }
    }

    /**
     * 死循环检测，校验相同集合不能重复遍历
     */
    public static class EndlessLoopException extends RuntimeException {
        public EndlessLoopException(String message) {
            super(message);
        }
    }

    /**
     * 死循环检测，校验相同集合不能重复遍历
     */
    public static class EndlessLoopCheckStack extends Stack<Context.Frame> {
        private final Set<Object> visitSet = Collections.newSetFromMap(new IdentityHashMap<>());

        @Override
        public Context.Frame push(Context.Frame item) {
//            if (item.data instanceof Collection || item.data instanceof Map) {
//                if (!visitSet.add(item.data)) {
//                    throw new EndlessLoopException(String.format("StackSize = %s, EndlessLoop at %s", size(), item.data));
//                }
//            }
            return super.push(item);
        }

        @Override
        public synchronized Context.Frame pop() {
            Context.Frame item = super.pop();
//            if (item.data instanceof Collection || item.data instanceof Map) {
//                visitSet.remove(item.data);
//            }
            return item;
        }
    }

    @Getter
    @Setter
    public static class Context extends BeanMap {
        private final Stack<Frame> renderDataStack = new EndlessLoopCheckStack();
        private int compileHit = 0;
        private int compileTotal = 0;
        private XWPFParagraph mergePos;
        private XWPFRun origMergePos;
        private List<RenderFinishRenderPolicy> finishRenderPolicyList = new ArrayList<>();

        public void dumpStack() {
            dumpStack(System.err);
        }

        public void dumpStack(PrintStream printStream) {
            int i = 0;
            for (Object o : renderDataStack) {
                printStream.printf("level=%d, data=%s\n", i, o);
                i++;
            }
        }

        public void incrementCompileHit() {
            compileHit++;
        }

        public void incrementCompileTotal() {
            compileTotal++;
        }

        public Object putTagRun(String tagName, ElementTemplate runTemplate) {
            return put(getTagRunKey(tagName), runTemplate);
        }

        public ElementTemplate getTagRun(String tagName) {
            return (ElementTemplate) get(getTagRunKey(tagName));
        }

        public int getOrigMergePosDocumentIndex() {
            if (origMergePos == null) {
                return -1;
            }
            IRunBody parent = origMergePos.getParent();
            XWPFDocument document = parent.getDocument();
            int bodyIndexByBody = findBodyIndexByBody(document, parent);
            return bodyIndexByBody;
        }

        public int getTagBodyIndex(IBody body, String tagName) {
            ElementTemplate tagRun = getTagRun(tagName);
            IRunBody parent = ((RunTemplate) tagRun).getRun().getParent();
            int bodyIndexByBody = findBodyIndexByBody(body, parent);

            if (bodyIndexByBody == -1) {
                bodyIndexByBody = findBodyIndexByTagName(body, tagRun.getSource());
            }

            if (bodyIndexByBody == -1) {
                for (Character symbol : COMPILE_CONFIGURE.getGramerChars()) {
                    String tagSource = buildSource(COMPILE_CONFIGURE, symbol, tagRun.getTagName());
                    bodyIndexByBody = findBodyIndexByTagName(body, tagSource);
                    if (bodyIndexByBody != -1) {
                        break;
                    }
                }
            }
            return bodyIndexByBody;
        }

        @AllArgsConstructor
        @Getter
        public static class Frame {
            private final ElementTemplate eleTemplate;
            private final Object data;
            private final XWPFTemplate template;

            @Override
            public String toString() {
                StringJoiner joiner = new StringJoiner(",");
                for (MetaTemplate elementTemplate : template.getElementTemplates()) {
                    joiner.add(elementTemplate.variable());
                }
                return (eleTemplate != null ? eleTemplate.getSource() : "null") + ":" + data + ". " + joiner;
            }
        }
    }

    /**
     * 全部渲染完后再渲染
     */
    public static abstract class RenderFinishRenderPolicy implements RenderPolicy {
        private boolean renderFinishFlag = false;

        @Override
        public final void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
            if (!renderFinishFlag) {
                return;
            }
            onFinishRender(eleTemplate, data, template);
        }

        public void onRenderFinish() {
            this.renderFinishFlag = true;
        }

        public abstract String tagName();

        public abstract void onFinishRender(ElementTemplate eleTemplate, Object data, XWPFTemplate template);
    }

    @Data
    public static class CustomField<NAME, CONTENT extends RenderData>
            extends BeanMap
            implements RenderData,
//            Ordered,
            AutowiredTemplate {
        public static final String FIELD_NAME = "name";
        public static final String FIELD_CONTENT = "content";
        public static final String FIELD_CHOSEROWDOCX = "choseRowDocx";
        /**
         * 字段名称
         */
        private NAME name;
        /**
         * 字段内容
         */
        private CONTENT content;
        private CustomStyle style;
        private byte[] docx;
        private int order;
        private XWPFTemplate template;
        private Function<CustomField<NAME, CONTENT>, byte[]> choseRowDocx;

        public CustomField() {
        }

        public CustomField(NAME name, CONTENT content) {
            this.name = name;
            this.content = content;
            this.style = CustomStyle.SINGLE_TEXT;
        }

        public static <NAME, CONTENT extends DocxProjectFileRenderData> CustomField<NAME, CONTENT> newTableRow(
                NAME name, CONTENT content, byte[] docx) {
            CustomField<NAME, CONTENT> customField = new CustomField<>();
            customField.setContent(content);
            customField.setName(name);
            customField.setStyle(CustomStyle.TABLE_ROW);
            customField.setDocx(docx);
            return customField;
        }

        public static <NAME extends TextRenderFilterData, CONTENT extends DocxProjectFileRenderData> CustomField<NAME, CONTENT> newTableTableRow(
                NAME name, CONTENT content, byte[] docx, Supplier<CustomField> supplier) {
            CustomField<NAME, CONTENT> customField = supplier.get();
            customField.setContent(content);
            customField.setName(name);
            customField.setStyle(CustomStyle.TABLE_TABLE_ROW);
            customField.setDocx(docx);
            return customField;
        }

        public boolean isEmptyCustomField() {
            CONTENT content = getContent();
            if (content instanceof DocxRenderData) {
                return ((DocxRenderData) content).getDataModels().isEmpty();
            } else {
                return DeepooveUtil.isEmpty(getContent());
            }
        }

        @Override
        public String toString() {
            if (name == null) {
                return "{" +
                        "content=" + content +
                        '}';
            } else {
                return "{" +
                        "name=" + name + "," +
                        "content=" + content +
                        '}';
            }
        }
    }

    /**
     * 跳过空数据
     *
     * @param <T>
     */
    public static class SkipEmptyValueList<T> extends LinkedList<T> {
        public SkipEmptyValueList(T v) {
            add(v);
        }

        public SkipEmptyValueList(Collection<T> v) {
            addAll(v);
        }

        public SkipEmptyValueList() {
        }

        public static void main(String[] args) {
            List list = new SkipEmptyValueList<>();
            list.addAll(Arrays.asList("1", 3, 4, 5));
            list.addAll(0, Arrays.asList(999, 388));
            System.out.println("list = " + list);
        }

        @Override
        public boolean add(T element) {
            if (DeepooveUtil.isEmpty((element))) {
                return false;
            }
            return super.add(element);
        }

        @Override
        public void add(int index, T element) {
            if (DeepooveUtil.isEmpty((element))) {
                return;
            }
            super.add(index, element);
        }

        @Override
        public boolean addAll(Collection<? extends T> c) {
            boolean b = false;
            for (T element : c) {
                b = add(element);
            }
            return b;
        }

        @Override
        public boolean addAll(int index, Collection<? extends T> c) {
            int i = 0;
            for (T element : c) {
                if (DeepooveUtil.isEmpty((element))) {
                    continue;
                }
                add(i, element);
                i++;
            }
            return i > 0;
        }

        public SkipEmptyValueList addAll() {
            addAll(new ArrayList<>(this));
            return this;
        }

        public SkipEmptyValueList clearAll() {
            super.clear();
            return this;
        }
    }

    /**
     * 多级文本
     */
    public static class MultistageTextRenderFilterData extends TextRenderFilterData {
        /**
         * N个空格等于一个tab
         */
        public static int WHITESPACE_TAB_COUNT = 4;
        @Getter
        @Setter
        private int level;
        @Getter
        @Setter
        private Map<Integer, String> levelMap;


        public MultistageTextRenderFilterData(String text, int levelOffset, Map<Integer, String> levelMap) {
            super(text);
            addLevel(levelOffset);
            this.levelMap = levelMap;
        }

        public MultistageTextRenderFilterData(String text, Style style, int levelOffset, Map<Integer, String> levelMap) {
            super(text, style);
            addLevel(levelOffset);
            this.levelMap = levelMap;
        }

        public static TextRenderFilterData newInstance(String value, Integer multistageOffset, Map<Integer, String> levelMap) {
            TextRenderFilterData renderFilterData;
            if (multistageOffset != null) {
                renderFilterData = new MultistageTextRenderFilterData(value, multistageOffset, levelMap);
            } else {
                renderFilterData = new TextRenderFilterData(value);
            }
            return renderFilterData;
        }

        public static TextRenderFilterData newInstance(String value, Style style, Integer multistageOffset, Map<Integer, String> levelMap) {
            TextRenderFilterData renderFilterData;
            if (multistageOffset != null) {
                renderFilterData = new MultistageTextRenderFilterData(value, style, multistageOffset, levelMap);
            } else {
                renderFilterData = new TextRenderFilterData(value, style);
            }
            return renderFilterData;
        }

        public static int countLevel(String string) {
            if (string == null || string.isEmpty()) {
                return 0;
            }
            int whitespaceCount = 0;
            int tabCount = 0;
            for (int i = 0; i < string.length(); i++) {
                char c = string.charAt(i);
                if (c == '\n' || c == '\r') {
                    //skip
                } else if (c == '\t') {
                    tabCount++;
                } else if (Character.isWhitespace(c)) {
//                    whitespaceCount++;
                } else {
                    break;
                }
            }
            //默认从1级开始, 1级,2级,3级
            return (whitespaceCount / WHITESPACE_TAB_COUNT) + tabCount + 1;
        }

        public static void main(String[] args) {
            int count1 = countLevel("    1");
            int count2 = countLevel("\t\t1");
            int count3 = countLevel("    \t\t1");
            int count4 = countLevel(" \t  \t \t1");
            System.out.println(" = ");
        }

        @Override
        public String filter(String string) {
            String filter = super.filter(string);
            this.level = countLevel(filter);
            return filter.trim();
        }

        public void addLevel(int add) {
            this.level = Math.max(level + add, 0);
        }
    }

    /**
     * 过滤特殊符号
     */
    public static class TextRenderFilterData extends TextRenderData implements AutowiredTemplate{
        @Getter
        @Setter
        private BiConsumer<RunTemplate, TextRenderFilterData> renderBeforeListener;
        @Getter
        @Setter
        private BiConsumer<RunTemplate, TextRenderFilterData> renderAfterListener;
        @Getter
        @Setter
        private XWPFTemplate template;
        @Getter
        @Setter
        private Consumer<DocElement> astDocConsumer = new ASTDocConsumer();
        /**
         * 首个符号 (请在标题字段, 也就自定义数据的 name字段上设置, 因为首个就是name字段. )
         */
        @Getter
        @Setter
        private String firstSymbol;

        public TextRenderFilterData() {
        }

        public TextRenderFilterData(String text) {
            super(text);
            setText(text);
        }

        public TextRenderFilterData(String color, String text) {
            super(color, text);
            setText(text);
        }

        public TextRenderFilterData(String text, Style style) {
            super(text, style);
            setText(text);
        }

        public static TextRenderFilterData wrap(String value) {
            return new TextRenderFilterData(value);
        }

        public static List<TextRenderFilterData> wrap(String[] values) {
            List<TextRenderFilterData> list = new SkipEmptyValueList<>();
            for (String value : values) {
                list.add(new TextRenderFilterData(value));
            }
            return list;
        }

        @Override
        public boolean equals(Object obj) {
            return super.equals(obj);
        }

        @Override
        public int hashCode() {
            return super.hashCode();
        }

        public String filter(String string) {
            return stringFilter(string);
        }

        @Override
        public String getText() {
            String text = super.getText();
            return text;
        }

        @Override
        public void setText(String text) {
            super.setText(filter(text));
        }

        public TextRenderFilterData bindRenderBeforeListener(BiConsumer<RunTemplate, TextRenderFilterData> renderBeforeListener) {
            this.renderBeforeListener = renderBeforeListener;
            return this;
        }

        public TextRenderFilterData bindRenderAfterListener(BiConsumer<RunTemplate, TextRenderFilterData> renderAfterListener) {
            this.renderAfterListener = renderAfterListener;
            return this;
        }

        @Override
        public String toString() {
            return getText();
        }
    }

    public static class DocxCustomFieldRenderData<NAME, CONTENT extends RenderData> extends DocxProjectFileRenderData<CustomField<NAME, CONTENT>> {
        public DocxCustomFieldRenderData(byte[] docx) {
            super(docx);
        }

        public DocxCustomFieldRenderData(byte[] docx, List<CustomField<NAME, CONTENT>> dataList) {
            super(docx, dataList);
        }

        public static DocxCustomFieldRenderData wrap(byte[] docx, Style style, String[] values) {
            List<CustomField> list = new SkipEmptyValueList<>();
            if (values != null) {
                for (String value : values) {
                    list.add(new CustomField<>("", new TextRenderFilterData(value, style)));
                }
            }
            return new DocxCustomFieldRenderData(docx, list);
        }

        @Override
        public List<CustomField<NAME, CONTENT>> getDataModels() {
            return super.getDataModels();
        }
    }

    /**
     * 会从项目中寻找文件资源
     */
    public static class DocxProjectFileRenderData<T> extends DocxRenderData {
        /**
         * 合并方法可重写. 需要返回合并后的文档, 如果返回文档为NULL, 则按照原始合并逻辑走.
         */
        @Getter
        @Setter
        private MergedOverrideMethod mergedOverrideMethod;
        /**
         * 渲染后的监听事件
         */
        @Getter
        @Setter
        private Consumer3<DocxProjectFileRenderData, XWPFTemplate,ElementTemplate> renderAfterListener;
        /**
         * 渲染前的监听事件
         */
        @Getter
        @Setter
        private Consumer3<DocxProjectFileRenderData, XWPFTemplate,ElementTemplate> renderBeforeListener;

        public DocxProjectFileRenderData(byte[] docx) {
            super(docx, null);
        }

        public DocxProjectFileRenderData(byte[] docx, List<?> dataList) {
            super(docx, dataList);
        }

        public static DocxProjectFileRenderData<CustomField<Object, TextRenderFilterData>> wrap(byte[] docx, Style style, String[] values, Integer multistageOffset, Map<Integer, String> levelMap) {
            return wrap(docx, style, values, false, multistageOffset, levelMap);
        }

        public static DocxProjectFileRenderData<CustomField<Object, TextRenderFilterData>> wrap(byte[] docx, Style style, String[] values, boolean upperCase, Integer multistageOffset, Map<Integer, String> levelMap) {
            List<Map> list = new SkipEmptyValueList<>();
            if (values != null) {
                for (String value : values) {
                    CustomField<Object, TextRenderFilterData> content = newContent(style, upperCase, value, multistageOffset, levelMap);
                    list.add(content);
                }
            }
            return new DocxProjectFileRenderData<>(docx, list);
        }

        public static CustomField<Object, TextRenderFilterData> newContent(Style style, boolean upperCase, String value, Integer multistageOffset, Map<Integer, String> levelMap) {
            CustomField<Object, TextRenderFilterData> map = new CustomField<>();
            if (upperCase) {
                value = toUpperCase(value);
            }
            TextRenderFilterData renderFilterData = MultistageTextRenderFilterData.newInstance(value, style, multistageOffset, levelMap);
            map.setContent(renderFilterData);
            return map;
        }

        @Override
        public byte[] getDocx() {
            return super.getDocx();
        }

        @Override
        public List<T> getDataModels() {
            return (List<T>) super.getDataModels();
        }

        public DocxProjectFileRenderData bindMergedOverrideMethod(MergedOverrideMethod mergedOverrideMethod) {
            setMergedOverrideMethod(mergedOverrideMethod);
            return this;
        }

        public DocxProjectFileRenderData bindRenderAfterListener(Consumer3<DocxProjectFileRenderData, XWPFTemplate,ElementTemplate> renderAfterListener) {
            setRenderAfterListener(renderAfterListener);
            return this;
        }

        public DocxProjectFileRenderData bindRenderBeforeListener(Consumer3<DocxProjectFileRenderData, XWPFTemplate,ElementTemplate> renderBeforeListener) {
            setRenderBeforeListener(renderBeforeListener);
            return this;
        }

        @Override
        public String toString() {
            StringJoiner joiner = new StringJoiner(",\n");
            int i = 0;
            for (Object o : getDataModels()) {
                Map map = BeanMap.toMap(o);
                Object name = map.get(CustomField.FIELD_NAME);
                Object content = map.get(CustomField.FIELD_CONTENT);
                if (name == null) {
                    joiner.add("{[" + i + "]content=" + content + "}");
                } else {
                    joiner.add("{[" + i + "]name=" + name + ",content=" + content + "}");
                }
                i++;
            }
            return joiner.toString();
        }
    }

    /**
     * 自定义字段渲染,则写入空文本或删除文档元素的渲染
     */
    public static class CustomFieldRenderPolicy extends DocxRenderPolicy implements RenderPolicy {
        /**
         * 重复样式
         *
         * @param destRun
         * @param srcRun
         */
        public static void styleRun(XWPFRun destRun, XWPFRun srcRun) {
            if (null == destRun || null == srcRun) {
                return;
            }
            destRun.setBold(srcRun.isBold());
//            destRun.setColor(srcRun.getColor());
            // destRun.setCharacterSpacing(srcRun.getCharacterSpacing());
            destRun.setFontFamily(srcRun.getFontFamily());
            int fontSize = srcRun.getFontSize();
            if (-1 != fontSize) {
                destRun.setFontSize(fontSize);
            }
            destRun.setItalic(srcRun.isItalic());
            destRun.setStrikeThrough(srcRun.isStrikeThrough());
            destRun.setUnderline(srcRun.getUnderline());
        }

        @Override
        public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
            onRenderBefore(eleTemplate, data, template);
            try {
                if (data instanceof DocxCustomFieldRenderData) {
                    renderByDocxCustomField(eleTemplate, (DocxCustomFieldRenderData) data, template);
                } else if (data instanceof DocxProjectFileRenderData) {
                    renderIfExistRender(eleTemplate, (DocxProjectFileRenderData) data, template);
                } else {
                    RenderPolicy policy = getPolicy(template.getConfig(), eleTemplate.getTagName(), com.deepoove.poi.config.GramerSymbol.TEXT.getSymbol());
                    if (policy != null) {
                        policy.render(eleTemplate, data, template);
                    }
                }
            } finally {
                onRenderAfter(eleTemplate, data, template);
            }
        }

        private void renderByDocxCustomField(ElementTemplate eleTemplate, DocxCustomFieldRenderData renderData, XWPFTemplate template) {
            Configure config = template.getConfig();
            NiceXWPFDocument doc = template.getXWPFDocument();
            RunTemplate runTemplate = (RunTemplate) eleTemplate;
            XWPFRun run = runTemplate.getRun();
            Consumer3<DocxProjectFileRenderData, XWPFTemplate, ElementTemplate> renderBeforeListener = renderData.getRenderBeforeListener();
            if (renderBeforeListener != null) {
                renderBeforeListener.accept(renderData, template, eleTemplate);
            }
            List<NiceXWPFDocument> docMerges = getMergedDocxs(eleTemplate,renderData, config, (t, o, docs, rawMethod) -> {
                CustomField customField = (CustomField) o;
                //枚举是给要特殊处理的渲染逻辑用, 如果要特殊处理, 就自己定义个枚举处理一下.
                switch (customField.getStyle()) {
                    case TABLE_ROW: {
                        renderByTableRow(runTemplate, t, customField);
                        break;
                    }
                    case TABLE_TABLE_ROW: {
                        renderByTableTableRow(runTemplate, t, customField, docs);
                        break;
                    }
                    case SINGLE_TEXT:
                    default: {
                        DeepooveUtil.render(t, customField);
                        break;
                    }
                }
                return t.getXWPFDocument();
            });

            if (!docMerges.isEmpty()) {
                clearText(run);
                try {
                    doc = tryMergedOverrideMethod(renderData, template, doc, run, docMerges, renderData.getMergedOverrideMethod());
                } catch (Exception e) {
                    log.error("merge docx error={}", e.toString(), e);
                }
                template.reload(doc);
            }
            Consumer3<DocxProjectFileRenderData, XWPFTemplate, ElementTemplate> renderAfterListener = renderData.getRenderAfterListener();
            if (renderAfterListener != null) {
                renderAfterListener.accept(renderData, template, eleTemplate);
            }
        }

        private XWPFTemplate renderByTableTableRow(RunTemplate customFieldRunTemplate, XWPFTemplate template, CustomField datas, List<NiceXWPFDocument> docMerges) throws Exception {
            Configure config = template.getConfig();
            XWPFDocument document = customFieldRunTemplate.getRun().getDocument();
//            List<ElementTemplate> elementTemplates = template.getElementTemplates();

            XWPFRun writeTableRowPos = addLine(document);
            XWPFTemplate newTemplate = DeepooveUtil.compile(datas.getDocx(), config, datas);
//            CONTEXT_LOCAL.get().setMergePos((XWPFParagraph) writeTableRowPos.getParent());

            DeepooveUtil.render(newTemplate, datas);
            NiceXWPFDocument mergeAfterDocument = merge(template.getXWPFDocument(),
                    Arrays.asList(newTemplate.getXWPFDocument()), writeTableRowPos);
            template.reload(mergeAfterDocument);
            return template;
        }

        private void renderByTableRow(RunTemplate customFieldRunTemplate, XWPFTemplate template, CustomField datas) {
            Configure config = template.getConfig();
            NiceXWPFDocument document = template.getXWPFDocument();
            List<ElementTemplate> elementTemplates = (List) template.getElementTemplates();

            RunTemplate nameRunTemplate = getRunTemplate(elementTemplates, CustomField.FIELD_NAME);
            Object nameRenderData = datas.getName();
            RenderPolicy namePolicy = getPolicy(config, nameRunTemplate.getTagName(), nameRunTemplate.getSign());

            RunTemplate contentRunTemplate = getRunTemplate(elementTemplates, CustomField.FIELD_CONTENT);
            Object contentRenderData = datas.getContent();
            RenderPolicy contentPolicy = getPolicy(config, contentRunTemplate.getTagName(), contentRunTemplate.getSign());

            //这俩用的是一个table
            XWPFRun rawNameRun = nameRunTemplate.getRun();
            XWPFRun rawContentRun = contentRunTemplate.getRun();

            XWPFRun customFieldNameRun = rawNameRun;
            XWPFRun customFieldContentRun = addLine(document);

            styleRun(customFieldContentRun, rawContentRun);

            nameRunTemplate.setRun(customFieldNameRun);
            contentRunTemplate.setRun(customFieldContentRun);

            namePolicy.render(nameRunTemplate, nameRenderData, template);
            contentPolicy.render(contentRunTemplate, contentRenderData, template);
        }

        /**
         * docx如数据的数组为空 则写入空文本或删除文档元素的渲染
         *
         * @param eleTemplate
         * @param docxRenderData
         * @param template
         */
        private void renderIfExistRender(ElementTemplate eleTemplate, DocxProjectFileRenderData docxRenderData, XWPFTemplate template) {
            NiceXWPFDocument doc = template.getXWPFDocument();
            RunTemplate runTemplate = (RunTemplate) eleTemplate;
            if (null == docxRenderData) {
                return;
            }
            Consumer3<DocxProjectFileRenderData, XWPFTemplate, ElementTemplate> renderBeforeListener = docxRenderData.getRenderBeforeListener();
            if (renderBeforeListener != null) {
                renderBeforeListener.accept(docxRenderData, template, eleTemplate);
            }
            XWPFRun run = runTemplate.getRun();
            List<NiceXWPFDocument> docMerges = getMergedDocxs(eleTemplate, docxRenderData, template.getConfig(), null);
            boolean isExistDocMerges = !docMerges.isEmpty();
            try {
                if (isExistDocMerges) {
                    clearText(run);
                    doc = tryMergedOverrideMethod(docxRenderData, template, doc, run, docMerges, docxRenderData.getMergedOverrideMethod());
                    template.reload(doc);
                } else {
                    int index = findBodyIndexByBody(doc, run.getParent());
                    //删掉元素,不仅仅是置为空
                    if (!removeBodyElement(doc, index, eleTemplate.getSource(), docxRenderData)) {
                        //没删掉就清数据. 没删掉说明有依赖关系,不能删
                        clearText(run);
                    }
                    template.getElementTemplates().remove(eleTemplate);
                }
                Consumer3<DocxProjectFileRenderData, XWPFTemplate,ElementTemplate>  renderAfterListener = docxRenderData.getRenderAfterListener();
                if (renderAfterListener != null) {
                    renderAfterListener.accept(docxRenderData, template, eleTemplate);
                }
            } catch (Exception e) {
                log.error("merge docx error", e);
            }
        }

        private NiceXWPFDocument tryMergedOverrideMethod(DocxRenderData renderData, XWPFTemplate template, NiceXWPFDocument doc, XWPFRun run, List<NiceXWPFDocument> docMerges, MergedOverrideMethod mergedOverrideMethod) throws Exception {
            Context context = CONTEXT_LOCAL.get();

            Callable<NiceXWPFDocument> rawMergedMethod = () -> merge(doc, docMerges, run);
            NiceXWPFDocument mergedOverrideDoc = null;
            if (mergedOverrideMethod != null) {
                context.setOrigMergePos(run);
                try {
                    mergedOverrideDoc = mergedOverrideMethod.apply(template, renderData, docMerges, rawMergedMethod);
                } finally {
                    context.setOrigMergePos(null);
                }
            }
            NiceXWPFDocument result;
            if (mergedOverrideDoc == null) {
                result = rawMergedMethod.call();
            } else {
                result = mergedOverrideDoc;
            }
            return result;
        }
    }

    /**
     * 删除行标签渲染
     */
    public static class RemoveLineTagRenderPolicy extends RenderFinishRenderPolicy {
        @Override
        public String tagName() {
            return "_removeLine";
        }

        @Override
        public void onFinishRender(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
            onRenderBefore(eleTemplate, data, template);
            try {
                NiceXWPFDocument document = template.getXWPFDocument();
                String tagSource = eleTemplate.getSource();
                int tagIndex = findBodyIndexByTagName(document, tagSource);

                int removeDataIndex = tagIndex - 1;
                boolean removeSuccess = false;
                if (removeDataIndex >= 0) {
                    IBodyElement removeBody = document.getBodyElements().get(removeDataIndex);
                    boolean needRemove = isNeedRemove(removeBody);
                    if (needRemove) {
                        removeSuccess = removeBodyElement(document, removeDataIndex, tagSource, data);
                    }
                }
                int removeLineTagIndex = removeSuccess ? tagIndex - 1 : tagIndex;
                removeBodyElement(document, removeLineTagIndex, tagSource, data);
            } finally {
                onRenderAfter(eleTemplate, data, template);
            }
        }
    }

    /**
     * 动态文本渲染 用于支持动态元素类型
     */
    public static class DynamicTextRenderPolicy implements RenderPolicy {
        @Override
        public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
            onRenderBefore(eleTemplate, data, template);
            TextRenderData textRenderData = null;
            RunTemplate runTemplate = (RunTemplate) eleTemplate;
            try {
                XWPFRun run = runTemplate.getRun();
                String tagName = runTemplate.getTagName();
                Configure config = template.getConfig();

                //如果字段没有数据, 就把表格删除了
                String tagSource = eleTemplate.getSource();
                if (isEmpty(data)) {
                    NiceXWPFDocument document = template.getXWPFDocument();
                    int index = findBodyIndexByTagName(document, tagSource);
                    if (!removeBodyElement(document, index, tagSource, data)) {//没删掉就清数据. 没删掉说明有依赖关系,不能删
                        clearText(run);
                    }
                    return;
                }

                //路由
                if (data instanceof TextRenderData) {
                    textRenderData = (TextRenderData) data;
                } else if (data instanceof RenderData) {
                    //根据数据寻找渲染的实现
                    Character symbol = getSymbol(data.getClass());
                    if (symbol != null) {
                        RunTemplate newTemplate = new RunTemplate();
                        newTemplate.setRun(run);
                        newTemplate.setSign(symbol);
                        newTemplate.setTagName(tagName);
                        newTemplate.setSource(buildSource(config, symbol, tagName));

                        RenderPolicy policy = getPolicy(config, tagName, symbol);
                        policy.render(newTemplate, data, template);
                        return;
                    }
                }
                if (textRenderData == null) {
                    textRenderData = new TextRenderData(data.toString());
                }
                if (textRenderData instanceof TextRenderFilterData) {
                    TextRenderFilterData textRenderFilterData = (TextRenderFilterData) textRenderData;
                    if (textRenderFilterData.getRenderBeforeListener() != null) {
                        textRenderFilterData.getRenderBeforeListener().accept(runTemplate, textRenderFilterData);
                    }
                }
                setText(textRenderData, run);
            } finally {
                if (textRenderData instanceof TextRenderFilterData) {
                    TextRenderFilterData textRenderFilterData = (TextRenderFilterData) textRenderData;
                    if (textRenderFilterData.getRenderAfterListener() != null) {
                        textRenderFilterData.getRenderAfterListener().accept(runTemplate, textRenderFilterData);
                    }
                }
                onRenderAfter(eleTemplate, data, template);
            }
        }

        private Character getSymbol(Class type) {
            for (Map.Entry<Class<? extends RenderData>, Character> entry : RENDER_DATA_SYMBOL_MAP.entrySet()) {
                if (entry.getKey().isAssignableFrom(type)) {
                    return entry.getValue();
                }
            }
            return null;
        }

        public boolean isDocx(Class renderDataType, Configure config, String tagName) {
            boolean isDocx;
            Character symbol = getSymbol(renderDataType);
            if (symbol != null) {
                RenderPolicy targetPolicy = getPolicy(config, tagName, symbol);
                isDocx = targetPolicy instanceof DocxRenderPolicy;
            } else {
                isDocx = false;
            }
            if (isDocx) {
                log.trace("isDocx=true. renderDataType={}", renderDataType);
            }
            return isDocx;
        }
    }

}
