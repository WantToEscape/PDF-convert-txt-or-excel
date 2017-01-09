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

Sub test()
    
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
                    Sleep 500
                    
                    Application.CutCopyMode = False
                    AppActivate strChromeTitle
                    DoEvents
                    SendKeys "^a"
                    Sleep 200
                    
                    AppActivate strChromeTitle
                    DoEvents
                    SendKeys "^c"
                    Sleep 200
                    
                    AppActivate strChromeTitle
                    DoEvents
                    SendKeys "^a"
                    Sleep 200
                    
                    AppActivate strChromeTitle
                    DoEvents
                    SendKeys "^c"
                    DoEvents
                    SendKeys "^c"
                    DoEvents
                    SendKeys "^c"
                    DoEvents
                    SendKeys "^c"
                    DoEvents
                    SendKeys "^c"
                    
                    AppActivate strChromeTitle
                    DoEvents
                    SendKeys "^c"
                    Sleep 200
                    
                    DoEvents
                    SendKeys "^c"
                    Sleep 200
                    
                    DoEvents

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

