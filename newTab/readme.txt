���ʱ����(����ʱҲ�����޸�):
ActiveTabBackColor:ѡ��tab�ı���ɫ
ActiveTabTextColor:ѡ��tab������ɫ
Align:tab�����ֵĶ��뷽ʽ
BackColor:�����ؼ��ı���ɫ
BorerLine:�Ƿ��б���
Font:����
HoverActiveTabTextColor:���ָ��ѡ�е�tab������ɫ
HoverTabTextColor:���ָ��δѡ��tab������ɫ
Style:tab������ʽ
TabBackColor:δѡ��tab�ı���ɫ
TabHeight:tab�߶�(>=200)
TabTextColor:δѡ��tab������ɫ
TabWidthMax:tab�����

����ʱ���Ժͷ�����
AddTab(key As String,Caption As String, Optional Width As Single = -1,
	Optional Image As String = "",Optional ToolTipText As String = "", 
	Optional preTabIndex As Long = -1)
��ӻ����һ��tab��key����Ϊ��ֵ��width������ʱΪ�Զ���Ӧ���
image��������Ϊimagelist��index��key,preTabIndex��ʾΪ���뵽�ĸ�tab�ĺ���
��������ã����Զ�׷�ӵ����ͬʱ��tab��Ϊѡ�кͿɼ�
ExistsTab(key as String)
�Ƿ���ڴ˹ؼ��ֵ�tab
ImageList:����ImageList
Index2Key(index as Long)
index��key��ת��
Key2Index(key as String)
key��index��ת��
RemoveAll
�Ƴ�����tab
RemoveTab(index)
��ָָ����tab�����indexΪ��ֵ�����Ƴ���index��tab,�����Ƴ�keyΪindex��tab
SelectTab:���ػ����õ�ǰѡ�е�tab�������tab���ɼ������Զ��ƶ�ֱ���ɼ�
TabCount:��ǰ����tab��
Tabs(index):����ָ����tab��index������index��key����
TabVisible(index)
ĳ��tab�Ƿ�ɼ���indexʹ��ͬ��


�¼���
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event tabChange(key As String)
Public Event SelectChang(previous As String, current As String)
Public Event Hover(key As String)
Public Event Click(key As String)
Public Event DblClick(key As String)

TAB����
Key           As String		�ؼ���
Index         As Long		����
Caption       As String		����
Width         As Single		���
ToolTipText   As String		������ʾ�ı�
Image         As String		ͼ��
Active        As Boolean	�Ƿ�ѡ��
Hover         As Boolean	�Ƿ������ָ

