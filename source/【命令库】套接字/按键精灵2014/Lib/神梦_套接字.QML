[General]
SyntaxVersion=2
MacroID=41562ddb-6bba-40ec-ad40-766c63f9e1b8
[Comment]

[Script]
'==================================================================================================
'【按键精灵也玩套接字、Sock、UDP】
'版本：v1.2
'更新：2021.08.17
'作者：神梦无痕
'ＱＱ：1042207232
'交流群：584781753
'=======================================[需要脚本定制可以找我]=======================================
'
'-----------------------
'//初始化套接字对象
'返回值，是否成功
Function Sock_初始化()
    Dim 系统目录, DLL路径, code, 当前版本, rep, Ver, ws
    If Not Sock_检测版本 Then 
        Set rep = New RegExp : 当前版本 = "1.2"
        rep.IgnoreCase = True
        rep.[Global] = True
        rep.Pattern = "\{v(\d+\.\d+)\}"
        Set xmlHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        xmlHttp.Open "Get", "https://360biji.com/note/view/650616", False
        xmlHttp.Send 
        Ver = xmlHttp.ResponseText
        If rep.test(Ver) Then 
            Ver = rep.Execute(Ver).Item(0).SubMatches.Item(0)
        Else 
            ver = 0
        End If
        If StrComp(Ver, 当前版本) = 1 Then 
            TracePrint "<font color=FF00FF>命令库有新的版本v" & Ver & "</font>"
            TracePrint "<font color=FF00FF>请点击链接：<a href='https://360biji.com/note/view/650616'  target='_blank'>https://360biji.com/note/view/650616</a></font>"
        Else 
            Sock_检测版本 = True 
        End If
        Set xmlHttp = Nothing
    End If
    系统目录 = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(1).Path
    Dim WMI : Set WMI = GetObject("winmgmts:\\.\root\cimv2")
    Dim col : Set col = WMI.ExecQuery("Select * from Win32_ComputerSystem")
    Dim Is64Bit : For Each objItem In col : Is64Bit = (InStr(objItem.SystemType, "x64") > 0) : Next
    If Is64Bit Then 系统目录 = Replace(系统目录, "System32", "SysWOW64", 1, 1, 1)
    DLL路径 = 系统目录 & "\MSWinsock.OCX"
    If Not CreateObject("Scripting.FileSystemObject").FileExists(DLL路径) Then 
        Dim xmlHttp : Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
        xmlHttp.Open "Get", "https://gitee.com/smwh/smos/raw/master/OCX/MSWinsock.ocx", False
        xmlHttp.Send 
        Dim ADO : Set ADO = CreateObject("Adodb.Stream")
        ADO.Type = 1 : ADO.Mode = 3 : ADO.Open  : ADO.Write xmlHttp.ResponseBody : ADO.SaveToFile DLL路径, 2 : ADO.Close 
        Set ADO = Nothing : Set xmlHttp = Nothing : Delay 2500
    End If 
    Call Createobject("Wscript.Shell").Run(系统目录 & "\regsvr32 " & DLL路径 & " /s", 0)
    code = _
    "Function CByte(ByVal Str)" & vbCrLf &_
    "    Dim MD, node, I, StrH" & vbCrLf &_
    "    Set MD = CreateObject(""Msxml2.DOMDocument"")" & vbCrLf &_
    "    Set node = MD.CreateElement(""binary"")" & vbCrLf &_
    "    node.DataType = ""bin.hex""" & vbCrLf &_
    "    If IsArray(Str) Then" & vbCrLf &_
    "       For i = 0 To UBound(Str)" & vbCrLf &_
    "           StrH = StrH & Right(""0"" + Hex(Str(i)), 2)" & vbCrLf &_
    "        Next" & vbCrLf &_
    "    Else" & vbCrLf &_
    "        For i = 1 To len(Str)" & vbCrLf &_
    "            StrH = StrH & Right(""0"" + Hex(Asc(mid(Str, i, 1))), 2)" & vbCrLf &_
    "        Next" & vbCrLf &_
    "    End If" & vbCrLf &_
    "    node.Text = StrH" & vbCrLf &_
    "    CByte = node.NodeTypedValue" & vbCrLf &_
    "    Set node = Nothing" & vbCrLf &_
    "    Set MD = Nothing " & vbCrLf &_
    "End Function" & vbCrLf &_
    "Function CStrB(ByVal ByteS)" & vbCrLf &_
    "    Dim ADO" & vbCrLf &_
    "    If VarType(ByteS) = 8209 Then " & vbCrLf &_
    "        Set ADO = CreateObject(""ADODB.Stream"")" & vbCrLf &_
    "        ADO.Type = 1" & vbCrLf &_
    "        ADO.Mode = 3" & vbCrLf &_
    "        ADO.Open" & vbCrLf &_
    "        ADO.Write ByteS" & vbCrLf &_
    "        ADO.Position = 0" & vbCrLf &_
    "        ADO.Type = 2" & vbCrLf &_
    "        ADO.Charset = ""GB2312""" & vbCrLf &_
    "        CStrB = ADO.ReadText" & vbCrLf &_
    "        ADO.Close" & vbCrLf &_
    "        Set ADO =nothing" & vbCrLf &_
    "    End If" & vbCrLf &_
    "End Function" & vbCrLf &_
    "Function ReadByte(ByVal Path)" & vbCrLf &_
    "    Dim ADO " & vbCrLf &_
    "    Set ADO = CreateObject(""ADODB.Stream"")" & vbCrLf &_
    "    ADO.Type = 1" & vbCrLf &_
    "    ADO.Open" & vbCrLf &_
    "    ADO.LoadFromFile Path" & vbCrLf &_
    "    ReadByte = ADO.Read" & vbCrLf &_
    "    ADO.Close" & vbCrLf &_
    "    Set ADO =nothing" & vbCrLf &_
    "End Function" & vbCrLf &_
    "Function SaveToFile(ByVal Path, ByVal ByteS)" & vbCrLf &_
    "    Dim ADO" & vbCrLf &_
    "    If VarType(ByteS) = 8209 Then " & vbCrLf &_
    "        Set ADO = CreateObject(""ADODB.Stream"")" & vbCrLf &_
    "        ADO.Type = 1" & vbCrLf &_
    "        ADO.Mode = 3" & vbCrLf &_
    "        ADO.Open" & vbCrLf &_
    "        ADO.Write ByteS" & vbCrLf &_
    "        ADO.SaveToFile Path, 2" & vbCrLf &_
    "        ADO.Close" & vbCrLf &_
    "        Set ADO = Nothing" & vbCrLf &_
    "    End If" & vbCrLf &_
    "End Function" & vbCrLf &_
    "Function LenByte(ByVal ByteS)" & vbCrLf &_
    "    Dim ADO" & vbCrLf &_
    "    If VarType(ByteS) = 8209 Then " & vbCrLf &_
    "        Set ADO=CreateObject(""ADODB.Stream"")" & vbCrLf &_
    "        ADO.Mode = 3" & vbCrLf &_
    "        ADO.Type = 1" & vbCrLf &_
    "        ADO.Open" & vbCrLf &_
    "        ADO.Write ByteS" & vbCrLf &_
    "        LenByte=ADO.Size" & vbCrLf &_
    "        ADO.Close " & vbCrLf &_
    "        Set ADO=nothing" & vbCrLf &_
    "    End If " & vbCrLf &_
    "End Function" & vbCrLf &_
    "Function MidByte(ByVal ByteS, ByVal Start, ByVal Length)" & vbCrLf &_
    "    Dim ADO" & vbCrLf &_
    "    If VarType(ByteS) = 8209 Then " & vbCrLf &_
    "        Set ADO = CreateObject(""ADODB.Stream"")" & vbCrLf &_
    "        ADO.Type = 1" & vbCrLf &_
    "        ADO.Open" & vbCrLf &_
    "        ADO.Write ByteS" & vbCrLf &_
    "        ADO.Position=Start" & vbCrLf &_
    "        MidByte = ADO.Read(Length)" & vbCrLf &_
    "        ADO.Close " & vbCrLf &_
    "        Set ADO =nothing" & vbCrLf &_
    "    End if " & vbCrLf &_
    "End Function" & vbCrLf &_
    "Function ByteAdd(ByVal ByteArr)" & vbCrLf &_
    "    'ByteArr为数组， Array(字节集1, 字节集2, 字节集3...)" & vbCrLf &_
    "    Dim ADO" & vbCrLf &_
    "    If VarType(ByteArr) = 8204 Then " & vbCrLf &_
    "        Set ADO = CreateObject(""ADODB.Stream"")" & vbCrLf &_
    "        ADO.Type = 1" & vbCrLf &_
    "        ADO.Open " & vbCrLf &_
    "        For i = 0 To UBound(ByteArr)" & vbCrLf &_
    "            If VarType(ByteArr(i)) = 8209 Then" & vbCrLf &_
    "                ADO.Write ByteArr(i)" & vbCrLf &_
    "            End If " & vbCrLf &_
    "        Next" & vbCrLf &_
    "        ADO.Position=0" & vbCrLf &_
    "        ByteAdd = ADO.Read" & vbCrLf &_
    "        ADO.Close " & vbCrLf &_
    "        Set ADO = nothing" & vbCrLf &_
    "    End If " & vbCrLf &_
    "End Function" & vbCrLf &_
    "Function HexToDec(ByVal sHex)" & vbCrLf &_
    "   Dim D,H,i,Ia" & vbCrLf &_
    "   D = 0" & vbCrLf &_
    "   H = UCase(sHex)" & vbCrLf &_
    "   For i = 1 To Len(H)" & vbCrLf &_
    "       Ia = Asc(Mid(H, i, 1)) - 48" & vbCrLf &_
    "       If Ia > 9 Then Ia = Ia - 7" & vbCrLf &_
    "       D = D * 16 + Ia" & vbCrLf &_
    "   Next" & vbCrLf &_
    "   HexToDec = D" & vbCrLf &_
    "End Function" & vbCrLf
    ExecuteGlobal code
    '为 MSWinsock.OCX 设置许可证
    Set ws = CreateObject("Wscript.Shell")
    ws.RegWrite "HKEY_CLASSES_ROOT\Licenses\2C49F800-C2DD-11CF-9AD6-0080C7E7B78D\", "mlrljgrlhltlngjlthrligklpkrhllglqlrk", "REG_SZ"
    Set ws = Nothing 
    Set WinSock = CreateObject("MSWinsock.Winsock")
    If TypeName(WinSock) = "Winsock" Then 
        WinSock.Protocol = 1 '创建UDP协议
        Sock_初始化 = True
    Else 
        Sock_初始化 = False 
    End If 
End Function
'-----------------------
'销毁套接字对象
Sub Sock_销毁()
    Set WinSock = Nothing
End Sub
'-----------------------
'//可以把一个端口号固定为本控件使用，使得别的应用程序不能再使用这个端口。
'参数1【必选】：整数型，端口（范围0~65535）
Sub 方法_绑定端口(端口)
    WinSock.Close
    WinSock.Bind 端口
End Sub
'-----------------------
'//当连接建立后，就可以调用发送数据，该方法只有一个参数，就是要发送的数据。
'参数1【必选】：字符串，远程IP
'参数2【必选】：整数型，远程端口（范围0~65535）
'参数3【必选】：变体型，数据
Sub 方法_发送数据(远程IP, 远程端口, 数据)
    WinSock.RemoteHost = 远程IP
    WinSock.RemotePort = 远程端口
    WinSock.SendData 数据
End Sub 
'-----------------------
'当本地计算机接收到远程计算机的数据时，数据存放在缓冲区中，该方法可以从缓冲区中取出数据
'参数1【必选】：整数型，超时时间，单位秒，0为不限时间
'返回值：数组，数组0=对方ip，数组1=数据内容；获取超时：数组下标=-1
Function 方法_接收数据(超时)
    Dim s, revip
    s = Now
    方法_接收数据 = Array()
    If IsEmpty(超时) Then 超时 = 0
    Do While 超时 = 0 Or Datediff("s", s, Now) <= 超时
        Delay 50
        If WinSock.BytesReceived > 0 Then
            revStr = ""
            WinSock.GetData revStr, vbString
            revip = WinSock.RemoteHostIp
            方法_接收数据 = Array(revip, revStr)
            Exit Do
        End If
    Loop
End Function
'-----------------------
'//当连接建立后，发送一个文件给对方
'参数1【必选】：字符串，远程IP
'参数2【必选】：整数型，远程端口（范围0~65535）
'参数3【必选】：字符串，文件路径
'返回值：逻辑型，是否成功
Function 方法_发送文件(远程IP, 远程端口, 文件路径)
    Dim ADO, Data, HZname
    方法_发送文件 = False 
    If CreateObject("Scripting.FileSystemObject").FileExists(文件路径) Then 
        HZname = CreateObject("Scripting.FileSystemObject").Getextensionname(文件路径)
        WinSock.RemoteHost = 远程IP
        WinSock.RemotePort = 远程端口
        Data = ReadByte(文件路径)
        Data = ByteAdd(Array(CByte(Left("{File}" + HZname + String(16, 0), 16)), Data))
        WinSock.SendData Data
        方法_发送文件 = True 
    Else 
        TracePrint "出错，要发送的文件不存在！"
    End If 
End Function
'-----------------------
'接收从远程发送过来的文件
'参数1【必选】：整数型，超时时间，单位秒，0为不限时间
'返回值：数组，数组0=对方ip，数组1=字节集数据；获取超时：数组下标=-1
Function 方法_接收文件(保存路径, 超时)
    Dim s, revip, ADO, ret, HZname, Data
    s = Now
    方法_接收文件 = Array()
    If IsEmpty(超时) Then 超时 = 0
    Do While 超时 = 0 Or Datediff("s", s, Now) <= 超时
        Delay 50
        If WinSock.BytesReceived > 0 Then
            revStr = ""
            WinSock.GetData revStr, vbArray+vbByte
            revip = WinSock.RemoteHostIp
            ret = Replace(CStrB(MidByte(revStr, 0, 16)), Chr(0), "")
            If Left(ret, 6) = "{File}" Then 
                If ret <> "{File}" Then 
                    HZname = Replace(ret, "{File}", ".")
                    If StrComp(Right(保存路径, Len(HZname)), HZname, 1) <> 0 Then 
                        FGPath_Temp = Split(保存路径, "\")
                        FGPath_Temp(UBound(FGPath_Temp)) = CreateObject("Scripting.FileSystemObject").GetbaseName(保存路径) & HZname
                        保存路径 = Join(FGPath_Temp, "\")
                    End If
                Else
                    FGPath_Temp = Split(保存路径, "\")
                    FGPath_Temp(UBound(FGPath_Temp)) = CreateObject("Scripting.FileSystemObject").GetbaseName(保存路径)
                    保存路径 = Join(FGPath_Temp, "\")
                End If
                Data = MidByte(revStr, 16, LenByte(revStr)-16)
                Call SaveToFile(保存路径, Data)
                方法_接收文件 = Array(revip, Data)
            Else
                TracePrint "出错，协议头格式不正确！"
            End If
            Exit Do
        End If
    Loop
End Function
'-----------------------
'关闭套接字对象，取消端口绑定
Sub 方法_关闭()
    WinSock.Close
End Sub
'-----------------------
'取得本地主机名
Function 属性_取本地主机名()
    属性_取本地主机名 = WinSock.LocalHostName
End Function
'-----------------------
'取得本地主机IP
Function 属性_取本地主机IP()
    属性_取本地主机IP = WinSock.LocalIP
End Function
'-----------------------
'取得本地端口
Function 属性_取本地端口()
    属性_取本地端口 = WinSock.LocalPort
End Function
'-----------------------
'取得远程IP
Function 属性_取远程IP()
    属性_取远程IP = WinSock.RemoteHost
End Function
'-----------------------
'取得远程端口
Function 属性_取远程端口()
    属性_取远程端口 = WinSock.RemotePort
End Function
'-----------------------
'//获取创建sock的对象状态
'返回值：整数型，状态
/*
    0 缺省的。关闭 
    1 打开 
    2 侦听 
    3 连接挂起 
    4 识别主机 
    5 已识别主机 
    6 正在连接 
    7 已连接 
    8 同级人员正在关闭连接 
    9 错误
*/
Function 属性_取状态()
    属性_取状态 = WinSock.State
End Function
'-----------------------
'//获取接收到的（当前在接收端缓冲区内的）数据的数量
Function 属性_取数据量()
    属性_取数据量 = WinSock.BytesReceived
End Function
'-----------------------
'//发送远程开机（网卡唤醒）魔术数据包
'参数1【必选】：字符串，远程网卡MAC
'参数2【必选】：字符串，远程IP
'参数3【必选】：整数型，远程端口（范围0~65535）
'返回值：逻辑型，是否成功
Function 应用_网卡唤醒(远程网卡MAC, 远程IP, 远程端口)
    Dim 魔术包
    If Sock_初始化() Then 
        魔术包 = "FFFFFFFFFFFF"
        For 16
            魔术包 = 魔术包 + Replace(Replace(Trim(远程网卡MAC), ":", ""), "-", "")
        Next
        Call 方法_发送数据(远程IP, 远程端口, 魔术包)
        应用_网卡唤醒 = True
    Else 
        应用_网卡唤醒 = False 
    End If 
End Function
'-----------------------
'//发送时钟同步数据包，并同步本机时间
'参数1【必选】：字符串，时钟服务器IP
/*
    阿里：ntp.aliyun.com 
    腾讯：time1.cloud.tencent.com 
    微软：time.windows.com 
    苹果：time.apple.com 
    谷歌：time.google.com 
    高通：time.izatcloud.net
    官方：ntp.ntsc.ac.cn
    清华：ntp.tuna.tsinghua.edu.cn
*/
'返回值：逻辑型，是否成功
Function 应用_时钟同步(时钟服务器IP)
    Dim 数据包, t, i, 时标(4), 校验, WshShell
    If Sock_初始化() Then 
        revData = ""
        If 时钟服务器IP = "" Then 时钟服务器IP = "ntp.aliyun.com"
        数据包 = CByte(Array(27, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 208, 175, 95, 245, 35, 215, 8, 0))
        时标(1) = DateDiff("s", "1900-1-1 0:0:0", DateAdd("h", -8, Now)) + 4294967296 '因计算果已经溢出。故用 +4294967296 求补获得正确结果
        WinSock.RemoteHost = 时钟服务器IP
        WinSock.RemotePort = 123
        WinSock.SendData 数据包
        t = Timer()
        Do
            Delay 50
            If WinSock.BytesReceived > 0 Then
                时标(4) = DateDiff("s", "1900-1-1 0:0:0", DateAdd("h", -8, Now)) + 4294967296 '因计算果已经溢出。故用 +4294967296 求补获得正确结果
                WinSock.GetData revData, vbArray + vbByte
                For i = 1 To 4
                    时标(0) = 时标(0) & Right("0" & Hex(AscB(MidB(revData, 24 + i, 1))), 2)
                    时标(2) = 时标(2) & Right("0" & Hex(AscB(MidB(revData, 32 + i, 1))), 2)
                    时标(3) = 时标(3) & Right("0" & Hex(AscB(MidB(revData, 40 + i, 1))), 2)
                Next
                时标(2) = HexToDec(时标(2))
                时标(3) = HexToDec(时标(3))
                校验 = HexToDec(时标(0))
                If 校验 = 3501154293 Then 
                    补偿值 = (时标(2) - 时标(1) + 时标(3) - 时标(4)) / 2
                    revData = DateAdd("s", 补偿值, Now)
                    ' 修改电脑时间
                    Set WshShell = CreateObject("wscript.Shell")
                    WshShell.Run "cmd.exe /c date " & FormatDateTime(revData, 2), 0
                    WshShell.Run "cmd.exe /c time " & FormatDateTime(revData, 3), 0
                    Set WshShell = Nothing
                End If
                Exit Do
            End If
        Loop Until Timer() - t > 30
        应用_时钟同步 = revData
    End If
End Function
/*〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓【更新历史】〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓
套接字v1.2 2021.08.17
\
|-- 新增 应用_时钟同步()
|
|
套接字v1.1 2017.06.11
\
|-- 修复 类未授权使用问题
|
|
套接字v1.0 2017.06.10
\
|-- 新增 Sock_初始化()
|-- 新增 Sock_销毁()
|-- 新增 方法_绑定端口()
|-- 新增 方法_发送数据()
|-- 新增 方法_发送文件()
|-- 新增 方法_关闭()
|-- 新增 方法_接收数据()
|-- 新增 方法_接收文件()
|-- 新增 属性_取本地端口()
|-- 新增 属性_取本地主机IP()
|-- 新增 属性_取本地主机名()
|-- 新增 属性_取数据量()
|-- 新增 属性_取远程IP()
|-- 新增 属性_取远程端口()
|-- 新增 属性_取状态()
|-- 新增 应用_网卡唤醒()
|
|
*/'〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓