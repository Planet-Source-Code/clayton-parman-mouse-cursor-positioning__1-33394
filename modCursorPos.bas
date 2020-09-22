Attribute VB_Name = "modCursorPos"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'     This project was written in, and formatted for, Courier New font.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'
'  Author: Clayton Parman
'
'    Date: 04-03-02
'
'    Desc: Procedure to move the Mouse cursor to a specified
'          coordinate in relation to a control on a form.
'
'   Notes: Operation is screen resolution independent.
'          Will work for any control that has an hwnd property.
'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Private Type RECT
   Left   As Long
   Top    As Long
   Right  As Long
   Bottom As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long


Public Sub MoveMouseCursorTo(Cntrl As Control, _
                    Optional EdgeInitials As String, _
                    Optional Down As Long, _
                    Optional Over As Long)

  'Moves the Mouse cursor to the Edge(s) or Center of the control.
  
  '"Cntl" is the name of the Control to move the mouse cursor to.
  
  'Optional "EdgeInitial" is:  T=Top, B=Bottom, L=Left, R=Right
  '      or  combinations of:  TL, TR, BL, BR  (Centered is the default).
  
  'Optional "Down" and "Over" provide an extra degree of positioning.
  'A minus  "Down"  or "Over"  will move "Up" or "Over Left".
  
   Dim Edge   As RECT
   Dim Vert   As Long
   Dim Horz   As Long
   
  'GetWindowRect returns the Controls Left-Right-Top-Bottom edges.
  
   GetWindowRect Cntrl.hwnd, Edge
  
   Select Case UCase(EdgeInitials)
   Case "T"         'Centered Horizontal at Top Edge of control
         Vert = Edge.Top + 1
         Horz = Edge.Left + ((Edge.Right - Edge.Left) / 2)
   Case "B"         'Centered Horizontal at Bottom Edge of control
         Vert = Edge.Bottom - 1
         Horz = Edge.Left + ((Edge.Right - Edge.Left) / 2)
   Case "L"         'Centered Vertical at Left Edge of control
         Vert = Edge.Top + ((Edge.Bottom - Edge.Top) / 2)
         Horz = Edge.Left + 1
   Case "R"         'Centered Vertical at Right Edge of control
         Vert = Edge.Top + ((Edge.Bottom - Edge.Top) / 2)
         Horz = Edge.Right - 1
   Case "LT", "TL"  'Top-Left corner of control
         Vert = Edge.Top + 1
         Horz = Edge.Left + 1
   Case "LB", "BL"  'Bottom-Left corner of control
         Vert = Edge.Bottom - 1
         Horz = Edge.Left + 1
   Case "RT", "TR"  'Top-Right corner of control
         Vert = Edge.Top + 1
         Horz = Edge.Right - 1
   Case "RB", "BR"  'Bottom-Right corner of control
         Vert = Edge.Bottom - 1
         Horz = Edge.Right - 1
   Case Else        'Center of control (default)
         Vert = Edge.Top + ((Edge.Bottom - Edge.Top) / 2)
         Horz = Edge.Left + ((Edge.Right - Edge.Left) / 2)
   End Select
  
   SetCursorPos Horz + Over, Vert + Down

End Sub

