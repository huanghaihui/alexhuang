---
layout: post
title: Outlook Add in Dev
permalink: /2014/12/outlookdev
---
&nbsp;&nbsp;&nbsp;&nbsp;最近做了Outlook的plugin，过程有点漫长，但是对于我这种WPF开发小白来说还是有点收获的。
<br>

###先主要介绍下插件的功能：

1. 要一个可复选的列表，通过过滤Category删选邮件。
2. 此列表可以随时删掉和增加。
3.  需要一个TagCloud，表示当前邮件的所有的Categories，且可进行继续的删选邮件。
4. 在当前选择的邮件中，可以随时更新此邮件的Category。</li>
上述的几个要求是比较高层次的概括，具体细节更加复杂，将在后续的文章进行介绍。

###下面先介绍下开发流程：
1. 环境选择：VS2010＋Outlook2010。我原来是用Outlook2013的，但是后来在import form region的时候出现了问题，还是建议用2010。
2. 项目创建：在VS中选择新建项目－Office－Outlook2010 add in。
3. 项目模块开发：分别进行了复选列表，列表增删，TagCloud，邮件Categories更新。


每一个模块的功能图片显示:
1. My Tags Panel(包含复选框，搜索框，TagCloud)
![Alt "Usercontrol"](/images/UserControl.png)
2. Add rabbitButton to control usercontrol(控制是否显示当前的Panel)
![Alt "Rabbit Button"](/images/DisplayOrHidden.png)
3. 邮件的Categories,并附带提示输入功能
![Alt "mailCategories"](/images/mailCategory.png)
