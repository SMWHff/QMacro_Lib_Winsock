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
MacroID=bb6daab9-670a-4f55-b8f3-34fb1a5032a9
Description=【神梦_套接字】网络同步时间
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


/* 【时钟服务器IP】

    阿里：ntp.aliyun.com 
    腾讯：time1.cloud.tencent.com 
    微软：time.windows.com 
    苹果：time.apple.com 
    谷歌：time.google.com 
    高通：time.izatcloud.net
    官方：ntp.ntsc.ac.cn
    清华：ntp.tuna.tsinghua.edu.cn
*/

返回值 = Lib.神梦_套接字.应用_时钟同步("ntp.aliyun.com")
TracePrint 返回值