VERSION 5.00
Begin VB.Form Viewlist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Fill List Viewer"
   ClientHeight    =   4545
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3810
   Icon            =   "Viewlist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete Current Item"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H00FF0000&
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Viewlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.ListIndex = -1 Then
        If MsgBox("You Do Not Have An Entry Selected", vbExclamation) = vbOK Then Exit Sub
        End If
        If MsgBox("Are You Sure You Want To Delete The Selected Entry?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Dim a As Integer
    a = List1.ListIndex
    List1.RemoveItem a
    If a = List1.ListCount Then
        List1.ListIndex = a - 1
    Else
        List1.ListIndex = a
    End If
    Close #1
    Open App.Path & "\" & "AF_list.dat" For Output As 1
    For i = 0 To List1.ListCount - 1
        Print #1, List1.List(i)
      Next i
      Close #1
      
End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
Close #1
Dim TempName As String, TempNumber As String
    Open App.Path & "\" & "AF_list.dat" For Input As 1
    On Error Resume Next
    Do Until EOF(1)
        Line Input #1, TempName
        List1.AddItem TempName
       Loop
       Close #1
       List1.ListIndex = 0
      
        
End Sub
