Attribute VB_Name = "mdlMakeArg"
'*************������������ģ��*************
'*******************************************
'*****Design By ������ô��ɪ & winzjjj*****
'*******************************************
'**************Who are you?***************
'******************************************
'***********You aren't a copyer!!!**********
'******************************************
'*********"NICE SELF DO IT!!......."**********
'*******************************************
Dim ArgsforLaunch() As String
Dim AIndex As Integer

'���һ���²���
Private Sub ArgAdd(ByVal arg As String)
    ReDim Preserve ArgsforLaunch(AIndex)
    ArgsforLaunch(AIndex) = arg
    AIndex = AIndex + 1
End Sub

'�������ú��������������ֱ�Ϊ �ڴ��С��native·�����Ƿ����֤�顢�Ƿ���Բ���
Public Function ArgSettingfor16x(ByVal MemoryMByte As Long, ByVal Username As String, ByVal McVersion As String, ByVal NativePath As String, ByVal AssetPath As String, Optional ByVal Gamedir As String = ".minecraft", Optional ByVal MCSession As String = "${auth_session}", Optional ByVal IIMC As Boolean = True, Optional ByVal IPD As Boolean = True, Optional ByVal accessToken As String, Optional ByVal id As String)
    On Error GoTo errline
    Erase ArgsforLaunch
    '���ж��ǲ��Ǿɷ�ʽ
    If Dir(App.Path & "\.minecraft\versions\" & McVersion & "\minecraft.jar") <> "" Then
        ArgSettingforOldMC McVersion, MemoryMByte, Username, , , , , , , "\.minecraft\versions\" & McVersion & "\natives"
        Exit Function
    End If
    
    
    ArgAdd "-Xmx" & CStr(MemoryMByte) & "m"   'Minecraft����ڴ�
    ArgAdd "-Dfml.ignoreInvalidMinecraftCertificates=" & IIf(IIMC, "true", "false")   '������Чminecraft֤��
    ArgAdd "-Dfml.ignorePatchDiscrepancies=" & IIf(IPD, "true", "false")    '���Բ�������
    ArgAdd "-Djava.library.path=""" & NativePath & """"   'Native·��
    ArgAdd mdlMakeArg.Makelibcmd(App.Path & "\.minecraft\versions\" & McVersion & "\" & McVersion & ".json")  '��·��
   
Open App.Path & "\.minecraft\versions\" & McVersion & "\" & McVersion & ".json" For Input As #1
Do Until EOF(1)
    Line Input #1, Lip
    NeiR = NeiR & vbCrLf & Lip
Loop
Close #1

temp1 = InStr(NeiR, Chr(34) & "mainClass" & Chr(34) & ": " & Chr(34))
temp2 = InStr(temp1 + 15, NeiR, Chr(34))
Mainclass = Mid(NeiR, temp1 + 14, temp2 - temp1 - 14) 'Mainclass

'�жϰ汾��1.6+����1.5-
On Error Resume Next
temp1 = InStr(McVersion, ".")
temp2 = InStr(temp1 + 1, McVersion, ".")
temp3 = Mid(McVersion, temp1 + 1)

If accessToken <> "" And id <> "" And Val(temp3) > 5.9 Or accessToken <> "" And id <> "" And McVersion Like "??w???" Then
'�������� ֻ��1.6+��
������ = Mainclass & _
                 " --username " & Username & _
                 " --version " & McVersion & _
                 " --gameDir " & Gamedir & _
                 " --assetsDir " & AssetPath & _
                 " --uuid " & id & _
                 " --assetIndex " & McVersion & _
                 " --userProperties {}" & _
                 " --userType Legacy" & _
                 " --accessToken " & accessToken
    ArgAdd ������

ElseIf Val(temp3) <= 5.9 Or McVersion Like "b*" Or McVersion Like "a*" Then '�ڶ���С�ڵ���5.9
������ = Mainclass & _
                 " --username " & Username & _
                 " " & MCSession & _
                 " --gameDir " & Gamedir & _
                 " --assetsDir " & AssetPath
    ArgAdd ������
ElseIf Val(temp3) > 5.9 Or McVersion Like "??w???" Then
������ = Mainclass & _
                 " --username " & Username & _
                 " --session " & MCSession & _
                 " --version " & McVersion & _
                 " --gameDir " & Gamedir & _
                 " --assetsDir " & AssetPath & _
                 " --uuid ${auth_uuid}" & _
                 " --assetIndex " & McVersion & _
                 " --userProperties {}" & _
                 " --userType legacy" & _
                 " --accessToken ${auth_access_token}" & _
                 " --tweakClass cpw.mods.fml.common.launcher.FMLTweaker"
    ArgAdd ������
Else
������ = Mainclass & _
                 " --username " & Username & _
                 " " & MCSession & _
                 " --gameDir " & Gamedir & _
                 " --assetsDir " & AssetPath
    ArgAdd ������
End If
Exit Function
errline:
    MsgBox "Json�ļ���ʧ��Minecraft ����ʧ�ܡ�", vbCritical
    End
End Function

Public Function ArgSetting()

End Function

'���argsforlaunch�е����в���
Public Function OutputArg4Command()
    Dim astr
    For Each astr In ArgsforLaunch
        OutputArg4Command = IIf(OutputArg4Command = "", astr, OutputArg4Command & " " & astr)
    Next
End Function

'���ɿ����
Public Function Makelibcmd(JsonFilePath As String)
On Error GoTo errline
    Dim JarFilePath As String
    Dim cpcommand$
    JarFilePath = Replace(JsonFilePath, ".json", ".jar")
    cpcommand = "-cp """  '����cp
    Close
    Open JsonFilePath For Input As #1
Do Until EOF(1)
    Line Input #1, Lip
    NeiR = NeiR & vbCrLf & Lip
Loop
Close #1
temp1 = 1
  Do
    temp1 = InStr(temp1 + 1, NeiR, """" & "name" & """" & ": ")
    If temp1 = 0 Then Exit Do
    temp2 = InStr(temp1, NeiR, "},")
    If temp2 = 0 Then
        temp2 = InStr(temp1, NeiR, "],")
    End If
    temp3 = Mid(NeiR, temp1 + 9, temp2 - temp1 - 9)
    temp4 = InStr(temp3, """")
    temp5 = Mid(temp3, 1, temp4 - 1)
    If InStr(temp5, ":") <> 0 Then '�����İ汾��һ����ð��
        '����д
        temp5x1 = InStr(temp5, ":")
        temp5x2 = Mid(temp5, 1, temp5x1 - 1)
        temp5x3 = Mid(temp5, temp5x1)
        temp5 = Replace(temp5x2, ".", "\") & temp5x3
        temp6 = Replace(temp5, ":", "\")
        temp7 = InStr(temp5, ":")
        temp8 = InStr(temp7 + 1, temp5, ":")
        temp9 = Mid(temp5, temp7 + 1, temp8 - temp7 - 1)
        temp10 = Mid(temp5, temp8 + 1)
        Cha64 = Is64bit
        If Cha64 = True Then temp11 = "64" Else temp11 = "32"
        '��������temp3����
        temp12 = InStr(temp3, Chr(34) & "windows" & Chr(34) & ": " & Chr(34))
        If temp12 <> 0 Then 'Ҫ����natives
            temp13 = InStr(temp12 + 13, temp3, Chr(34))
            temp14 = Mid(temp3, temp12 + 12, temp13 - temp12 - 12)
            temp14 = Replace(temp14, "${arch}", temp11)
            temp15 = App.Path & "\.minecraft\libraries\" & temp6 & "\" & temp9 & "-" & temp10 & "-" & temp14 & ".jar"
            mul = App.Path & "\.minecraft\libraries\" & temp6 & "\"
            temp16 = Dir(mul)
            If temp16 <> "" Then
            cpcommand = cpcommand & temp15 & ";"
            End If
        End If
        temp15 = App.Path & "\.minecraft\libraries\" & temp6 & "\" & temp9 & "-" & temp10 & ".jar"
        luj = App.Path & "\.minecraft\libraries\" & temp15
        mul = App.Path & "\.minecraft\libraries\" & temp6 & "\"
        temp16 = Dir(mul)
        If temp16 <> "" Then
        cpcommand = cpcommand & temp15 & ";"
        End If
    End If
Loop
cpcommand = Left(cpcommand, Len(cpcommand) - 1)
cpcommand = cpcommand & ";" & JarFilePath & """"
Makelibcmd = cpcommand
    Open App.Path & "\libs.txt" For Output As #1
    Print #1, cpcommand
    Close #1
Exit Function
errline:
If Err.Number = 79 Then
    MsgBox "δ�ҵ��汾 Json,Minecraft ����ʧ��.", vbCritical
End If
End Function


Public Function ArgSettingforOldMC(ByVal McVersion As String, ByVal MemoryMByte As Long, ByVal Username As String, Optional ByVal PermSize As Long = 64, Optional ByVal MaxPermSize As Long = 128, Optional ByVal TWODNoddraw As String = "true", Optional ByVal TWODpmoffscreen As String = "false", Optional ByVal TWODd3d = "false", Optional ByVal TWODOpenGL = "false", Optional ByVal NativesPath = "\.minecraft\bin\natives")
    Erase ArgsforLaunch
    ArgAdd "-Xincgc"
    ArgAdd "-Xmx" & MemoryMByte & "M"
    ArgAdd "-XX:PermSize=" & PermSize & "M"
    ArgAdd "-XX:MaxPermSize=" & MaxPermSize & "M"
    ArgAdd "-Dsun.java2d.noddraw=" & TWODNoddraw
    ArgAdd "-Dsun.java2d.pmoffscreen=" & TWODpmoffscreen
    ArgAdd "-Dsun.java2d.d3d=" & TWODd3d
    ArgAdd "-Dsun.java2d.opengl=" & TWODOpenGL
    ArgAdd "-cp"
   ' Dim ss(), i
   ' ReDim ss(0)
   Dim sFiles() As String
    mdlFiles.TreeSearch App.Path & "\.minecraft\versions\" & McVersion & "\", "*.jar", sFiles()
    Dim cps$
    cps = """"
    For I = LBound(sFiles) To UBound(sFiles)
        cps = cps & sFiles(I) & ";"
    Next
    cps = Left(cps, Len(cps) - 1)
    cps = cps & """"
    ArgAdd cps
    ArgAdd "-Djava.library.path=""" & App.Path & NativesPath & """"
    ArgAdd "net.minecraft.client.Minecraft"
    ArgAdd Username
End Function

