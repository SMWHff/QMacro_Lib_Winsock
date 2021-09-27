[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=b5d9192d-f268-4ab2-9a78-fd5d1f2e0c28
Description=【神梦_套接字】通过网卡远程唤醒电脑
Enable=1
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
'#================================================================
'#        【命令库】套接字例子
'#----------------------------------------------------------------
'#        【作者】：神梦无痕
'#        【ＱＱ】：1042207232
'#        【Ｑ群】：624655641
'#        【更新】：2021-09-27
'#----------------------------------------------------------------
'#  插件说明：按键精灵也玩套接字、Sock、UDP
'#----------------------------------------------------------------
'#  神梦工具: http://pan.baidu.com/s/1dESHf8X
'#----------------------------------------------------------------
'#  特别声明：请勿用于非法用途！！（否则后果自负）
'#================================================================




//【惊喜】只要有路由器，便可实现远程唤醒电脑！
//http://bbs.anjian.com/showtopic-564454-1.aspx



//网卡唤醒电脑
远程网卡MAC = "AA:BB:CC:DD:EE:FF"
远程IP = "192.168.1.100"
远程端口 = 100

返回值 = Lib.神梦_套接字.应用_网卡唤醒(远程网卡MAC, 远程IP, 远程端口)
If 返回值 Then 
	TracePrint "发送成功！"
Else 
	TracePrint "发送失败！"
End If