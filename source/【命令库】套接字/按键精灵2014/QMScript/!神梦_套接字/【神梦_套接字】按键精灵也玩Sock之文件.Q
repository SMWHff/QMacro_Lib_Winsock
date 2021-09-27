[General]
MacroID=031401e9-d874-44b8-a9f6-bc6025a9d911
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
Description=【神梦_套接字】按键精灵也玩Sock之文件
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



//启动服务端
BeginThread 服务端
Delay 1000


//脚本开始
If Lib.神梦_套接字.Sock_初始化() Then
	
    '绑定用来获取信息的端口
    Call Lib.神梦_套接字.方法_绑定端口(1002)
	
	
    '给服务端发送一个文件
    Path = "C:\123.jpg"
    TracePrint "【客户端】发送：传输文件 "& Path &" 给服务器！"
    Call Lib.神梦_套接字.方法_发送文件("127.0.0.1", 1001, Path)
    Delay 3500
	
    '获取服务端反馈
    ret = Lib.神梦_套接字.方法_接收数据(0)
    Delay 3500
    If UBound(ret) > -1 Then
        TracePrint "【客户端】收到：" & ret(0) & ":" & ret(1)
    Else
        TracePrint "超出时间限制，未收到答案"
    End If
	
	'取消绑定
	Call Lib.神梦_套接字.方法_关闭()
End If
Call Lib.神梦_套接字.Sock_销毁()



'==========================【用另一个线程创建一个服务端】==========================
Sub 服务端()
    If Lib.神梦_套接字.Sock_初始化() Then
		
        '绑定用来获取信息的端口
        Call Lib.神梦_套接字.方法_绑定端口(1001)
		
		
        '接收客户端的请求
        Path = "C:\999.jpg"
        ret = Lib.神梦_套接字.方法_接收文件(Path,0)
        Delay 3500
        If UBound(ret) > - 1 Then
            TracePrint "【服务端】收到:" & ret(0) & ":保存文件：" & Path
            Delay 3500
			
            '接收成功后，反馈信息给客户端
            反馈 = "我是服务端，收到了你的文件！"
            TracePrint "【服务端】发送：" & 反馈
            Call Lib.神梦_套接字.方法_发送数据(ret(0), 1002, 反馈)
        End If
		
        '取消绑定
        Call Lib.神梦_套接字.方法_关闭()
    End If
    Call Lib.神梦_套接字.Sock_销毁()
End Sub