Attribute VB_Name = "main"

' Connecting the report to SAP and screenshotting all the projects (SOX Greece)
' These lines are necessary for SAP GUI automation
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

' These lines are responsible for simulating mouse movements
' Depending on your operating system, use these lines or the commented-out alternatives below

'Private Declare Sub mouse_event Lib "user32" _
'(ByVal dwFlags As Long, ByVal dx As Long, _
'ByVal dy As Long, ByVal cButtons As Long, _
'ByVal swextrainfo As Long)

Private Declare PtrSafe Sub mouse_event Lib "user32" _
(ByVal dwFlags As Long, ByVal dx As Long, _
ByVal dy As Long, ByVal cButtons As Long, _
ByVal swextrainfo As Long)

' Constants for simulating mouse button clicks
Private Const mouseeventf_leftdown = &H2
Private Const mouseeventf_leftup = &H4
Private Const mouseeventF_Rightdown As Long = &H8
Private Const mouseeventF_rightup As Long = &H10

' Depending on your operating system, declare functions with or without PtrSafe
' Public Declare Function SetCursorPos Lib "user32.dll" _
' (ByVal x As Integer, ByVal y As Integer) As Long

Public Declare PtrSafe Function SetCursorPos Lib "user32.dll" _
(ByVal x As Integer, ByVal y As Integer) As Long

' Access the sleep function to introduce delays between actions
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' Access the GetCursorPos function to retrieve current cursor position
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' Main procedure for handling the entire screenshot process for SOX Greece projects
Sub main_POC()

    ' Ask user if they want to proceed with the macro
    CarryOn = MSGBOX("Do you want to run this macro?", vbYesNo)
    If CarryOn = vbYes Then

        ' Set up SAP GUI session
        Dim sapGuiApp As Object
        Dim excelApp As Object
        Dim excelSheet As Object
        Dim session As Object
        Set session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

        ' Retrieve mouse coordinates stored in the worksheet (B4 and B5)
        Dim Hold As POINTAPI
        Hold.X_Pos = Workbooks("Greece screens Projects.xlsm").Worksheets("Macro").Range("B4").Value
        Hold.Y_Pos = Workbooks("Greece screens Projects.xlsm").Worksheets("Macro").Range("B5").Value

        ' Copy the project numbers from the POC worksheet to paste into SAP
        Workbooks("Greece screens Projects.xlsm").Worksheets("POC").Activate
        Range("A4").Select
        Range(Selection, Selection.End(xlDown)).Copy

        ' Interact with SAP to navigate to the appropriate fields and enter data
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/ctxt$6-KOKRS").Text = "9000"  ' Cost center
        session.findById("wnd[0]/usr/ctxt$6-KSTAR").Text = "PSR_NET"  ' Internal order
        session.findById("wnd[0]/usr/ctxt$6-KSTAR").SetFocus
        session.findById("wnd[0]/usr/ctxt$6-KSTAR").CaretPosition = 7
        session.findById("wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH").Press  ' Pressing button to proceed
        session.findById("wnd[1]/tbar[0]/btn[16]").Press  ' Confirmation button
        session.findById("wnd[1]/tbar[0]/btn[24]").Press
        session.findById("wnd[1]/tbar[0]/btn[8]").Press
        session.findById("wnd[0]/tbar[1]/btn[8]").Press
        session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").ExpandNode "000001"  ' Expand relevant node
        session.findById("wnd[0]/tbar[1]/btn[24]").Press  ' Press next button

        ' Click the triangle in SAP based on saved coordinates
        SetCursorPos Hold.X_Pos, Hold.Y_Pos
        ' Simulate mouse click
        mouse_event mouseeventf_leftdown, 0&, 0&, 0&, 0&
        ' Optional: Uncomment to simulate mouse release
        ' mouse_event mouseeventf_leftup, 0&, 0&, 0&, 0&
        Sleep (250)  ' Pause for 250 milliseconds

        ' Capture screenshot and paste it into Excel
        wh = Worksheets("Macro").Range("B7").Value  ' Retrieve width
        ht = Worksheets("Macro").Range("B8").Value  ' Retrieve height

        Application.SendKeys "({1068})", True  ' Send keystrokes for PrintScreen
        DoEvents
        Sleep (500)  ' Small delay
        Workbooks("Greece screens Projects.xlsm").Worksheets("POC").Activate
        Sleep (1000)  ' Wait for Excel to be ready
        Sheets("POC").Paste Destination:=Sheets("POC").Range("H8")  ' Paste screenshot

        ' Adjust the screenshot dimensions after pasting
        With ActiveSheet
            Set shp = .Shapes(.Shapes.Count)
        End With

        ' Crop the screenshot to the specified width and height
        h = -(ht - shp.Height)
        w = -(wh - shp.Width)

        shp.LockAspectRatio = False
        shp.PictureFormat.CropRight = w
        shp.PictureFormat.CropBottom = h
        shp.Select
        Selection.ShapeRange.ScaleWidth 0.6, msoFalse, msoScaleFromTopLeft
        Selection.ShapeRange.ScaleHeight 0.6, msoFalse, msoScaleFromTopLeft

        ' Loop through each project node and take further screenshots
        Dim Nodes() As Variant
        Nodes = Range("A4").CurrentRegion
        Dim i As Long
        
        For i = LBound(Nodes) To UBound(Nodes)
            If Nodes(i, 3) = "select" Then
                ' Select relevant project node in SAP
                session.findById("wnd[0]/shellcont/shell/shellcont[2]/shell").SelectedNode = Nodes(i, 2)
                session.findById("wnd[0]/tbar[1]/btn[24]").Press  ' Press next button

                ' Click triangle in SAP again
                SetCursorPos Hold.X_Pos, Hold.Y_Pos
                mouse_event mouseeventf_leftdown, 0&, 0&, 0&, 0&
                Sleep (250)

                ' Capture additional screenshots and paste them in the next available position in Excel
                Application.SendKeys "({1068})", True
                DoEvents
                Sleep (500)
                Windows("Greece screens Projects.xlsm").Activate
                DoEvents
                Sleep (600)
                Sheets("POC").Paste Destination:=Sheets("POC").Cells(50 + (i - 1) * 30, "H")

                ' Adjust dimensions of the new screenshot
                With ActiveSheet
                    Set shp = .Shapes(.Shapes.Count)
                End With

                h = -(ht - shp.Height)
                w = -(wh - shp.Width)

                shp.LockAspectRatio = False
                shp.PictureFormat.CropRight = w
                shp.PictureFormat.CropBottom = h
                shp.Select
                Selection.ShapeRange.ScaleWidth 0.55, msoFalse, msoScaleFromTopLeft
                Selection.ShapeRange.ScaleHeight 0.55, msoFalse, msoScaleFromTopLeft
            End If
        Next i

        ' Final steps: maximize SAP window and confirm closing actions
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/btn[3]").Press  ' Press close
        session.findById("wnd[1]/usr/btnBUTTON_YES").Press  ' Confirm action

    End If

    ' Notify user that the process is complete
    MSGBOX "Done"

End Sub



    
    
    
    
    




