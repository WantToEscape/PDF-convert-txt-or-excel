Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Private Const WM_CHAR = &H102             '  PostMessage String

Private Const BM_CLICK = &HF5             '  hWnd click

Private Const MOUSEEVENTF_LEFTDOWN = &H2  '  left button down

Private Const MOUSEEVENTF_LEFTUP = &H4    '  left button up


'---------------------------------------------
'Clear Office Clipboard
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As UUID, ppvObject As Object) As Long
Private Declare Function AccessibleChildren Lib "oleacc" (ByVal paccContainer As IAccessible, ByVal iChildStart As Long, ByVal cChildren As Long, rgvarChildren As Variant, pcObtained As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Type UUID
    lData1 As Long
    nData2 As Integer
    nData3 As Integer
    abytData4(0 To 7) As Byte
End Type

Private Const ROLE_PUSHBUTTON = &H2B&
Sub ClearOfficeClipboard()
    Dim hMain As Long
    Dim hExcel2 As Long
    Dim hClip As Long
    Dim hWindow As Long
    Dim hParent As Long
    Dim lParameter As Long
    Dim octl As CommandBarControl
    Dim oIA As IAccessible
    Dim oNewIA As IAccessible
    Dim tg As UUID
    Dim lReturn As Long
    Dim lStart As Long
    Dim avKids() As Variant
    Dim avMoreKids() As Variant
    Dim lHowMany As Long
    Dim lGotHowMany As Long
    Dim bClip As Boolean
    Dim i As Long
    Dim hVersion As Long
        
    hMain = Application.hwnd
    
    hVersion = Application.Version
    
    
    '用于取得剪切板窗口的句柄（剪切板窗口可见时）
    Do
        hExcel2 = FindWindowEx(hMain, hExcel2, "EXCEL2", vbNullString)
        hParent = hExcel2: hWindow = 0
        hWindow = FindWindowEx(hParent, hWindow, "MsoCommandBar", vbNullString)
        If hWindow Then
            hParent = hWindow: hWindow = 0
            hWindow = FindWindowEx(hParent, hWindow, "MsoWorkPane", vbNullString)
            If hWindow Then
                hParent = hWindow: hWindow = 0
                hClip = FindWindowEx(hParent, hWindow, "bosa_sdm_XL9", "Collect and Paste 2.0")
                If hClip > 0 Then Exit Do
            End If
        End If
    Loop While hExcel2 > 0
    
    
    
    If hClip = 0 Then
        With Application.CommandBars("Task Pane")
            If Not .Visible Then
                LockWindowUpdate hMain
                Set octl = Application.CommandBars(1).FindControl(ID:=809, recursive:=True)
                If Not octl Is Nothing Then octl.Execute
                .Visible = False
                LockWindowUpdate 0
            End If
        End With
        hParent = hMain: hWindow = 0
        hWindow = FindWindowEx(hParent, hWindow, "MsoWorkPane", vbNullString)
        If hWindow Then
            hParent = hWindow: hWindow = 0
            hClip = FindWindowEx(hParent, hWindow, "bosa_sdm_XL9", "Collect and Paste 2.0")
        End If
    End If
    '假如Excel版本为2007版且剪切板不可见时使其可见
    If hVersion = 12 Then
        bClip = True
        With Application.CommandBars("Office Clipboard")
            If Not .Visible Then
                LockWindowUpdate hMain
                bClip = False
                Set octl = Application.CommandBars(1).FindControl(ID:=809, recursive:=True)
                If Not octl Is Nothing Then octl.Execute
            End If
        End With
    End If
    
    '用于取得剪切板窗口的句柄（剪切板窗口可见时）
    If hClip = 0 Then
        Do
            hExcel2 = FindWindowEx(hMain, hExcel2, "EXCEL2", vbNullString)
            hParent = hExcel2: hWindow = 0
            hWindow = FindWindowEx(hParent, hWindow, "MsoCommandBar", vbNullString)
            If hWindow Then
                hParent = hWindow: hWindow = 0
                hWindow = FindWindowEx(hParent, hWindow, "MsoWorkPane", vbNullString)
                If hWindow Then
                    hParent = hWindow: hWindow = 0
                    hClip = FindWindowEx(hParent, hWindow, "bosa_sdm_XL9", "Collect and Paste 2.0")
                    If hClip > 0 Then Exit Do
                End If
            End If
        Loop While hExcel2 > 0
    End If
    
    '即如以上都未找到剪切板窗口，显示错误信息
    If hClip = 0 Then
        MsgBox "剪切板窗口未找到"
        Exit Sub
    End If
    
    '以下部分代码参考了《Advanced Microsoft Visual Basic 6.0 Second Edition》第16章Microsoft Active Accessibility部分
    '定义IAccessible对象的GUID{618736E0-3C3D-11CF-810C-00AA00389B71}
    With tg
        .lData1 = &H618736E0
        .nData2 = &H3C3D
        .nData3 = &H11CF
        .abytData4(0) = &H81
        .abytData4(1) = &HC
        .abytData4(2) = &H0
        .abytData4(3) = &HAA
        .abytData4(4) = &H0
        .abytData4(5) = &H38
        .abytData4(6) = &H9B
        .abytData4(7) = &H71
    End With
    '从窗体返回Accessible对象
    lReturn = AccessibleObjectFromWindow(hClip, 0, tg, oIA)
    lStart = 0
    '/取得Accessible的子对象数量
    lHowMany = oIA.accChildCount
    ReDim avKids(lHowMany - 1) As Variant
    lGotHowMany = 0
    '/返回Accessible的子对象
    lReturn = AccessibleChildren(oIA, lStart, lHowMany, avKids(0), lGotHowMany)
    For i = 0 To lGotHowMany - 1
        If IsObject(avKids(i)) = True Then
            If avKids(i).accName = "Collect and Paste 2.0" Then
                Set oNewIA = avKids(i)
                lHowMany = oNewIA.accChildCount
                Exit For
            End If
        End If
    Next i
    ReDim avMoreKids(lHowMany - 1) As Variant
    lReturn = AccessibleChildren(oNewIA, lStart, lHowMany, avMoreKids(0), lGotHowMany)
    '取得"全部清空"按钮并执行它
    For i = 0 To lHowMany - 1
        If IsObject(avMoreKids(i)) = False Then
            If oNewIA.accName(avMoreKids(i)) = "Clear All" And oNewIA.accRole(avMoreKids(i)) = ROLE_PUSHBUTTON Then
                oNewIA.accDoDefaultAction (avMoreKids(i))
                'Exit For
            ElseIf oNewIA.accName(avMoreKids(i)) = "" Then
                oNewIA.accDoDefaultAction (avMoreKids(i))
            
            End If
        End If
    Next i
    
    
    
    'Application.CutCopyMode = False
    
End Sub

Function GetClipboardText() As String
    Dim MyData As DataObject, MyStr As String
    Set MyData = New DataObject
    MyData.GetFromClipboard    '获得剪切板内容
    GetClipboardText = MyData.GetText     '赋值给变量
End Function
'---------------------------------------------

Sub Main()
    
    Dim i%
    Application.ScreenUpdating = False
    
    
    'Check if the chrome window is open.
    Dim getChrome As Boolean
    Dim strExcelTitle$, strChromeTitle$
        For Each Process In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='Chrome.exe'")
            getChrome = True
            Exit For
        Next
        If getChrome = False Then
            MsgBox "Can't find the chrome , pls open the Google Chrome."
            Exit Sub
        End If
        strExcelTitle = Trim(Split(Application.Caption, "-")(0))
        strChromeTitle = "chrome"
        
    

    'Select pdf folder.
    Dim Filepath$
    Dim strPath$, myPath$, arrPath
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = -1 Then
                Filepath = .SelectedItems(1) & "\"
            Else
                MsgBox "Pls select a folder !", , "End"
                Exit Sub
            End If
        End With
        myPath = Dir(Filepath & "*.pdf")
        If myPath = "" Then
            MsgBox "Can't find pdf in this filepath.", , "End"
            Exit Sub
        Else
            strPath = myPath
            Do
                myPath = Dir
                If myPath = "" Then Exit Do
                strPath = strPath & "|" & myPath
            Loop
        End If
        arrPath = Split(strPath, "|")

    
    
    'Use the chrome browser to open the PDF and use sendkey to copy/paste pdf data.
    Dim myWB As Workbook
    Dim sheetCount%
    Dim hWnd As Long, hWndEdit As Long, hWndOpen As Long
    Dim waitTime%
    Dim x As Long, y As Long
    Dim mySheet As Worksheet, sheetName As String
        Set myWB = Workbooks.Add
        sheetCount = myWB.Sheets.Count
        If sheetCount > 1 Then
            Application.DisplayAlerts = False
            For i = sheetCount To 2 Step -1
                myWB.Sheets(i).Delete
            Next i
            Application.DisplayAlerts = True
        End If
        
        
        
        x = GetSystemMetrics32(0) / 2   ' Screen Width center
        y = GetSystemMetrics32(1) / 2   ' Screen Height center
        
        
        AppActivate strChromeTitle
        SendKeys "^t"
        Sleep 1000
    DoEvents
        For i = 0 To UBound(arrPath)
            AppActivate strChromeTitle
            
            Sleep 100
            
            hWnd = 0: waitTime = 0
            Sleep 100
            SendKeys "^o"
            Sleep 200
            
            Do
                DoEvents
                Sleep 200
                hWnd = FindWindow(vbNullString, "Open")
                If hWnd <> 0 Then
                    hWndEdit = FindWindowEx(hWnd, 0, "ComboBoxEx32", vbNullString)
                    hWndEdit = FindWindowEx(hWndEdit, 0, "ComboBox", vbNullString)
                    hWndEdit = FindWindowEx(hWndEdit, 0, "Edit", vbNullString)
                    hWndOpen = FindWindowEx(hWnd, 0, "Button", "&Open")
                    
                    myPath = Filepath & arrPath(i)
                    For j = 1 To VBA.Len(myPath)
                        StrData = VBA.Mid(myPath, j, 1)
                        PostMessage hWndEdit, WM_CHAR, Asc(StrData), 0
                    Next j
                    PostMessage hWndOpen, BM_CLICK, 0, 0
                    Sleep 500
                    
                    AppActivate strChromeTitle
                    Call SetCursorPos(x, y)
                    mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
                    Sleep 20
                    mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
                    Sleep 200
                    
                    Call ClearOfficeClipboard
    
    
                    Do    'Until not the clipboard is nothing 
                        AppActivate pidChrome
                            DoEvents
                        SendKeys "^a"
                            Sleep 200
                        SendKeys "^c", True
                            DoEvents
                            Sleep 200
                        SendKeys "^c", True
                            DoEvents
                            Sleep 200
                        Err = 0
                        On Error Resume Next
                            strClipboard = GetClipboardText
                            If Err <> -2147221404 Then
                                Exit Do
                            End If
                            Call SetCursorPos(x, y)
                            mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
                            Sleep 20
                            mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
                            Sleep 200
                        On Error GoTo 0
                    Loop

                    'Creat new sheet for this pdf
                    sheetName = Replace(arrPath(i), ".pdf", "")
                    If Len(sheetName) > 30 Then
                        sheetName = Left(sheetName, 30)
                    End If
                    Call AddSheet(sheetName)
                    Set mySheet = Sheets(sheetName)
                    mySheet.Activate
                    Cells(1, 1).Select
                    ActiveSheet.PasteSpecial Format:="Unicode Text", link:=False, DisplayAsIcon:=False
                    
                    Sleep 500
                    Exit Do
                End If
                waitTime = waitTime + 1
                If waitTime Mod 20 = 0 Then   'have been waited 4s
                    SendKeys "^o"
                    Sleep 200
                ElseIf waitTime > 50 Then
                    MsgBox "Error,Waited time is too long"
                    Exit Sub
                End If
            Loop
            'Exit For
        Next i
    
    MsgBox "Done!"
    'AppActivate "chrome"
    
    
    'AppActivate "Microsoft Excel"


End Sub


Private Sub AddSheet(sheetName As String)      'add new sheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        On Error Resume Next
        If ActiveWorkbook.Sheets(sheetName) Is Nothing Then
            ActiveWorkbook.Worksheets.Add().Name = sheetName
        Else
            Sheets(sheetName).Delete
            Call AddSheet(sheetName)
        End If
    Sheets(sheetName).Cells.NumberFormat = "@"
End Sub

