VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypewriterEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Listens to Word so typewriter mode can keep the cursor at the end.
Option Explicit

Public WithEvents App As Word.Application

' Any click, arrow key or selection snaps the cursor back to the end.
Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    If Typewriter.LocksDocument(Sel.Document) Then
        Typewriter.CheckTimeUp
        Typewriter.PinCursor
    End If
End Sub

Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    Typewriter.ReleaseDocument Doc
End Sub
