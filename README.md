# poi-report (用poi-tl实现)

### 简介

1.推荐报告
我们平时内推需要给一个简历和大概描述一下候选人，
猎头也一样，不过他提供的是要体现猎头专业性的一份PDF报告，叫《推荐报告》。
为了体现专业性，统一报告格式，降低制作成本。在系统里有个功能叫制作推荐报告。

这个推荐报告随着产品需求的增加，令这个能力愈来越支持个性化定制。
为了不与业务参杂一起，这个技术与业务剥离开。

例： 富文本加粗，按一个tab键是圆圈或对勾，2个tab键是梅花，3个tab等等。。
富文本在特定情况下回车认为是段落，不是换行。 页面多模块嵌套。解决用户因多敲回车导致的模块间距增加（不能让用户敲上的东西影响美观）。