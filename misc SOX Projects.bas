Attribute VB_Name = "misc"

' These lines are necessary for SAP GUI integration
Public SapGuiAuto, WScript, msgcol
Public objGui  As GuiApplication
Public objConn As GuiConnection
Public session As GuiSession

' Constants and declarations for screen metrics and mouse actions
Public Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
Public Const SM_CXSCREEN = 0 ' Screen width
Public Const SM_CYSCREEN = 1 ' Screen height

' Declaration for simulating mouse events
Private Declare PtrSafe Sub mouse_event Lib "user32" _
(ByVal dwFlags As Long, ByVal dx As Long, _
ByVal dy As Long, ByVal cButtons As Long, _
ByVal swextrainfo As Long)

' Constants for mouse events (left and right button clicks)
Private Const mouseeventf_leftdown = &H2
Private Const mouseeventf_leftup = &H4
Private Const mouseeventF_Rightdown As Long = &H8
Private Const mouseeventF_rightup As Long = &H10

' Declaration for setting cursor position on the screen
Public Declare PtrSafe Function SetCursorPos Lib "user32.dll" _
(ByVal x As Integer, ByVal y As Integer) As Long

' Declaration for accessing the sleep function to introduce delays
Public Declare PtrSafe Sub Sleep Lib "kernel32" Alias "sleep" (ByVal dwMilliseconds As Long)

' Declaration for retrieving the current cursor position
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' Custom data type to hold the X and Y coordinates for cursor position
Type POINTAPI
    X_Pos As Long
    Y_Pos As Long
End Type

' Main routine to dimension variables, retrieve cursor position,
' and display coordinates in the worksheet
Sub Get_Cursor_Pos()

    ' Run the macro and give the user 3 seconds to position the cursor
    Application.Wait (Now() + TimeValue("0:00:03"))
    
    ' Dimension the variable that will hold the x and y cursor positions
    Dim Hold As POINTAPI
    
    ' Retrieve the current cursor positions and store them in the Hold variable
    GetCursorPos Hold

    ' Display the cursor position coordinates to the user
    MSGBOX "X Position is : " & Hold.X_Pos & Chr(10) & "Y Position is : " & Hold.Y_Pos
    
    ' Store the cursor coordinates in the worksheet for later use
    Worksheets("Macro").Range("B4") = Hold.X_Pos
    Worksheets("Macro").Range("B5") = Hold.Y_Pos
    
    ' Optional: Uncomment the line below to set the cursor position back to where it was
    ' SetCursorPos Hold.X_Pos, Hold.Y_Pos
End Sub

' Routine to check the screen resolution of a screenshot and crop it accordingly
Sub Check_resolution()

    ' Retrieve the width and height stored in the worksheet
    wh = Worksheets("Macro").Range("B7").Value
    ht = Worksheets("Macro").Range("B8").Value

    ' Send the Print Screen command to take a screenshot
    Application.SendKeys "({1068})", True
    DoEvents
    
    ' Paste the screenshot into the Macro worksheet
    Sheets("Macro").Paste Destination:=Sheets("Macro").Range("A20")

    ' Select the last shape (the screenshot) and crop it based on the provided dimensions
    With ActiveSheet
        Set shp = .Shapes(.Shapes.Count)
    End With

    ' Adjust the crop dimensions
    h = -(ht - shp.Height)
    w = -(wh - shp.Width)

    ' Crop the image to the correct dimensions
    shp.LockAspectRatio = False
    shp.PictureFormat.CropRight = w
    shp.PictureFormat.CropBottom = h

End Sub

' Routine to check the coordinates and simulate mouse clicks at those positions
Sub Check_coordinates()

    ' Dimension the POINTAPI variable to hold cursor coordinates
    Dim Hold As POINTAPI
    
    ' Initialize SAP GUI scripting objects
    Set SapGuiAuto = GetObject("SAPGUI")
    Set objGui = SapGuiAuto.GetScriptingEngine
    Set objConn = objGui.Children(0)
    Set session = objConn.Children(0)
    
    ' Maximize the SAP window
    session.findById("wnd[0]").maximize
    
    ' Retrieve previously stored cursor coordinates from the worksheet
    Hold.X_Pos = Worksheets("Macro").Range("B4").Value
    Hold.Y_Pos = Worksheets("Macro").Range("B5").Value
    
    ' Set the cursor position to the stored coordinates
    SetCursorPos Hold.X_Pos, Hold.Y_Pos
    
    ' Simulate a left mouse button click at the cursor position
    mouse_event mouseeventf_leftdown, 0&, 0&, 0&, 0&
    mouse_event mouseeventf_leftup, 0&, 0&, 0&, 0&

End Sub

' Routine to clean shapes from all specified sheets
Sub Cleaning_sheets()

    ' Activate the POC worksheet and delete all shapes (screenshots)
    Workbooks("Greece screens Projects.xlsm").Worksheets("POC").Activate
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    
    ' Repeat for the other specified worksheets
    Workbooks("Greece screens Projects.xlsm").Worksheets("CCM").Activate
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    
    Workbooks("Greece screens Projects.xlsm").Worksheets("CCM Service").Activate
    ActiveSheet.Shapes.SelectAll
    Selection.Delete
    
    Workbooks("Greece screens Projects.xlsm").Worksheets("WAR").Activate
    ActiveSheet.Shapes.SelectAll
    Selection.Delete

    ' Notify the user that the cleaning process is complete
    MSGBOX "Done"

End Sub




