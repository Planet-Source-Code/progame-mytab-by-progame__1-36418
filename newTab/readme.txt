设计时属性(运行时也可以修改):
ActiveTabBackColor:选中tab的背景色
ActiveTabTextColor:选中tab的文字色
Align:tab内文字的对齐方式
BackColor:整个控件的背景色
BorerLine:是否有边线
Font:字体
HoverActiveTabTextColor:鼠标指向选中的tab的文字色
HoverTabTextColor:鼠标指向未选中tab的文字色
Style:tab放置样式
TabBackColor:未选中tab的背景色
TabHeight:tab高度(>=200)
TabTextColor:未选中tab的文字色
TabWidthMax:tab最大宽度

运行时属性和方法：
AddTab(key As String,Caption As String, Optional Width As Single = -1,
	Optional Image As String = "",Optional ToolTipText As String = "", 
	Optional preTabIndex As Long = -1)
添加或插入一个tab，key不能为数值，width不设置时为自动适应宽度
image可以设置为imagelist的index或key,preTabIndex表示为插入到哪个tab的后面
如果不设置，则自动追加到最后，同时此tab将为选中和可见
ExistsTab(key as String)
是否存在此关键字的tab
ImageList:关联ImageList
Index2Key(index as Long)
index到key的转换
Key2Index(key as String)
key到index的转换
RemoveAll
移除所有tab
RemoveTab(index)
移指指定的tab，如果index为数值，则移除第index个tab,否则移除key为index的tab
SelectTab:返回或设置当前选中的tab，如果此tab不可见，将自动移动直至可见
TabCount:当前所有tab数
Tabs(index):引用指定的tab，index可以是index和key引用
TabVisible(index)
某个tab是否可见，index使用同上


事件：
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event tabChange(key As String)
Public Event SelectChang(previous As String, current As String)
Public Event Hover(key As String)
Public Event Click(key As String)
Public Event DblClick(key As String)

TAB对象：
Key           As String		关键字
Index         As Long		索引
Caption       As String		名称
Width         As Single		宽度
ToolTipText   As String		工具提示文本
Image         As String		图像
Active        As Boolean	是否选中
Hover         As Boolean	是否鼠标所指

