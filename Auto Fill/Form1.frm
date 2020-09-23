VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto Fill Example"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Af_list 
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "View Currently Recodred Lists"
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6495
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":1CFA
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Fill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   39
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   2040
         TabIndex        =   11
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Text :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add current word to Auto Fill List"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Special Thanks : Harry, Guru, E.Murphy, M.Jackson, Harish,  "
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":25C4
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LB_FINDSTRING = &H18F
Private Const BACKSPACE = 8
Private Const DELETE = 46

Private Declare Function SendMessage Lib "user32" Alias _
      "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParm As Long, lParm As Any) As Long
      

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
        MsgBox "You Must Type Atleast something to add. In the meantime why not vote for me, it hardly takes a minute.", vbExclamation, ":( Entry Not Added"
Else
Dim i
Af_list.AddItem Text1.Text
On Error Resume Next
Open App.Path & "\" & "AF_list.dat" For Output As 1
        For i = 0 To Af_list.ListCount - 1
        Print #1, Af_list.List(i)
               Next i
        Close #1
        MsgBox "Your Entry Has BeenM Added. Please VOTE FOR ME", vbInformation, ":( Entry Added"
End If

End Sub

Private Sub Command3_Click()
frmAbout.Show
End Sub

Private Sub Command4_Click()
Viewlist.Show
Unload Me
End Sub

Private Sub Form_Load()
'load the auto fill list
    Dim TempName As String, TempNumber As String
    Open App.Path & "\" & "AF_list.dat" For Input As 1
    On Error Resume Next
    Do Until EOF(1)
        Line Input #1, TempName
        Af_list.AddItem TempName
        Loop
    Close #1
    Af_list.ListIndex = 0
End Sub

Private Sub Text1_Change()
  Dim Location As Long
  
                  ' Send a message to Windows to find the string in the
                  ' text box
  Af_list.ListIndex = SendMessage(Af_list.hwnd, _
        LB_FINDSTRING, -1, ByVal CStr(Text1.Text))
        
  If Af_list.ListIndex <> -1 Then  ' Did we find a match?
    Location = Text1.SelStart           ' Yes, we did, so show
    Text1.Text = Af_list   ' what we have thus
    Text1.SelStart = Location           ' far
    Text1.SelLength = Len(Text1) - Location
  End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim temp As String
  
  If KeyCode = BACKSPACE Then               ' If there's a backspace...
    If Text1.SelLength <> 0 Then    ' ...and we have a string...
      Text1.Text = Mid$(Text1.Text, 1, Text1.SelStart - 1)
      KeyCode = 0
    End If
  ElseIf KeyCode = DELETE Then              ' Was it a DEL code?
                                            ' Do we have a string?
    If Text1.SelLength <> 0 And _
       Text1.SelStart <> 0 Then     ' Then clear all of it
       Text1.Text = ""
       KeyCode = 0
    End If
  End If
End Sub
