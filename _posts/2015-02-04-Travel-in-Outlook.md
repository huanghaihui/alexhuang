---
layout: post
title: Outlook Dev-TaskPane
permalink: /2015/02/04/outlookdev2
---
&nbsp;&nbsp;&nbsp;&nbsp;在这篇中，介绍一下WPF中的UserControl。这个控件可以讲多个控件作为子元素来嵌套其内。我们的目的是要将Categories的复选框，搜索框以及TagCloud放在一起，Usercontrol很符合我们的要求。

###邮件Categories的复选框设计
&nbsp;&nbsp;&nbsp;&nbsp;我选择了checkedListBox这个控件。选择这个控件的好处有以下几点：1.无需考虑每个check选项的位置，可以省去Draw 控件的时间。2.我们可以通过键盘输入字母，控件会自动跳到以该字母为首字母的选项，使用变得更加简洁方便。
<br>     如何处理这个Categories的复选框呢？首先得获取相关得Categories：当然Outlook有相应的对象去存储整个邮箱的Categories。
<code>
     Outlook.Categories categories =  Globals.ThisAddIn.Application.Session.Categories;
     </code>
然后就是将获取的categories存放的checkedListBox中。相关用法都可以在msdn上查到，这里就不做过多的介绍。
####保存选中的选项
这通过触发该控件的ItemCheck事件：<code> checkedListBox1.ItemCheck += new ItemCheckEventHandler(this.Check_Clicked);</code>
在Check_Clicked:<code>private void Check_Clicked(Object sender, ItemCheckEventArgs e)</code>方法中，我们可以通过<code> checkedListBox1.CheckedItems</code>接口获得所选中的checkItems，但是这里需要注意的是，当我们新选择一个Item时，结果并不包含这个Item，此时你就需要利用<code>e.NewValue</code>来判断。

####对选中的选项进行邮件搜索
这里先粗略的讲述下搜索方法，我们使用的是Outlook的自带搜
索框。利用了接口：<code>explorer.Search(searchTxt, Outlook.OlSearchScope.olSearchScopeAllFolders);</code>。我们是要进行categories的搜索，所以需要设置serachTxt 的格式：<code>searchTxt += "Category:=\"" + check + "\" ";</code>。这样我们就可以得到相应的视图结果。
![Alt "insideSearchBox"](/images/insideSearchBox.png)


####通过选中的邮件更新TagCloud
我们刚才的操作得出了我们所想要的一些邮件。与此同时，我们希望现在出现的邮件中的所有Categories能够出现在我们的CloudPane中，方便我们进一步删选。<br>
1.我们选用richtextbox作为基础控件，存放过滤后的category。至于为什么选这个，出自msdn上的一个案例，没有深究。
<br>
2.通过 <code> explorer.CurrentView</code>获取当前视图的邮件，但是这里需要转多个弯得到最终的邮件项，我会另外写一篇博客介绍其相关内容。
<br>
3.对于每一个获取的邮件项，我们通过找到他的category来保存。
<pre><code>Outlook.Row row = table.GetNextRow();
string entryID = row["EntryID"];
object item = nsp.GetItemFromID(entryID);
if (item is Outlook.MailItem)
{
  Outlook.MailItem mailitem = item as Outlook.MailItem;
  if (mailitem.Categories != null)
  update_TagClodeList(mailitem.Categories);
}
</code></pre>
因为每个邮件项中我们只能得到其特殊的EntryID，所以我们可以通过EntryID来找到邮件空间中相应的邮件，然后将其的categories更新到richtextbox中。
![Alt "updateColud"](/images/updateCloud.png)


####搜索框刷新checkedListBox控件
这里的目的是希望能够通过在搜索框的输入来进行刷新更新。这个做法比较简单，我们获取搜索框中的内容，然后对check列表进行扫描，包含搜索框中字段的则保留，不然清除。
<pre><code>String text = this.searchBox.Text;
shownList.Clear();
shownList.AddRange(itemList);
foreach (string check in itemList)
{
  if (!textContains(check, text))
  {
    shownList.Remove(check);
  }
  else if(!itemList.Contains(check))
  {
    shownList.Add(check);
  }
}
var items = new List<string>();</code></pre><pre><code>
foreach (string item in shownList) items.Add(item);
shownList.Clear();
shownList.AddRange(items.OrderBy(i => i).ToArray());
refreshChecker();</code></pre>

我们首先将存留的选项保存到list中，然后更新输出list。
<pre><code>private void refreshChecker()
{
  checkedListBox1.Items.Clear();
  foreach (string item in shownList)
  checkedListBox1.Items.Add(item,
  CheckState.Unchecked);
}</code></pre>
这样就可以保持搜索框及时更新了。
![Alt "searchBox"](/images/searchBox.png)
