VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHourglass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private lOldMousePtr As Long

Private Sub Class_Initialize()
   '## Show the Hourglass
   lOldMousePtr = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   Debug.Print "MousePointer::Hourglass"
End Sub

Private Sub Class_Terminate()
   '## Show previous pointer
   Screen.MousePointer = lOldMousePtr
   Debug.Print "MousePointer::Reset"
End Sub
