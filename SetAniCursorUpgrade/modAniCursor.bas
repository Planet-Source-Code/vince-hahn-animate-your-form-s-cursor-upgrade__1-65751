Attribute VB_Name = "modAniCursor"
Option Explicit

'This Module Was Made To Help Simplify The Code Found At

'     http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=24150&lngWId=1

' Cursor Icon was tweaked from someone else's icon I can't really say the name because I don't know...
' If anyone knows the icon's creator or at least who it is let me know please so I can give credit for it.  it was free online.


Public Declare Function LoadCursorFromFile Lib "user32" _
    Alias "LoadCursorFromFileA" _
    (ByVal lpFileName As String) As Long


Public Declare Function SetClassLong Lib "user32" _
    Alias "SetClassLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Const GCL_HCURSOR = (-12)

Dim sCursorFile As String
Dim hCursor As Long
Dim hOldCursor As Long
Dim lReturn As Long



Public Sub SetAniCursor(frmForm As Form, strLocation As String)

    ' Assign The Location of the file
    sCursorFile = strLocation
    
    'call cursor load api, pass the location
    hCursor = LoadCursorFromFile(sCursorFile)
    
    'call setclasslong declaration,passing the calling form's hwnd, gcl_hcursor constant,
    'and of course the value returned from the loadcursorfromfile api
    hOldCursor = SetClassLong(frmForm.hwnd, GCL_HCURSOR, hCursor)

End Sub

Public Sub KillAniCursor(frmForm As Form)

    'set lReturn to remove the, um,  "registered" i guess yo u can call it animated cursor, pass the value of hOldCursor's "registrar" data to
    'setclasslong api along with yet again the calling form's hwnd.
    lReturn = SetClassLong(frmForm.hwnd, GCL_HCURSOR, hOldCursor)

End Sub

