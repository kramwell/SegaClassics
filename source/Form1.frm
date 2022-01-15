VERSION 5.00
Begin VB.Form SegaClassics 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SegaClassics v1.1 "
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   ">> Random Game <<"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GitHub Repo"
      Height          =   375
      Left            =   3000
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go to Game"
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
      Begin VB.TextBox txtGoto 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdGoto 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      Height          =   375
      Left            =   1320
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   4545
      ItemData        =   "Form1.frx":08CA
      Left            =   0
      List            =   "Form1.frx":12AF
      TabIndex        =   1
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblGameno 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3855
      TabIndex        =   3
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Selected Game No."
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   1740
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   -240
      Picture         =   "Form1.frx":5995
      Top             =   -600
      Width           =   3375
   End
End
Attribute VB_Name = "SegaClassics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'v1   10/NOV/2006
'v1.1 15/JAN/2022
'SegaClassics v1.1 is an updated version of the much loved v1 edition with a few new tweaks!

Option Explicit
        
        'declarations for the shell command
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long


Private Declare Function FindExecutable Lib "shell32.dll" Alias _
         "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
         String, ByVal lpResult As String) As Long
       'ends the shell command

Private Sub cmdAbout_Click()
frmAbout.Show
End Sub

Private Sub cmdGoto_Click()
Dim intNumgoto As Integer
Dim intNumstore As Integer

On Error GoTo errorrepair

intNumstore = 843

'this says if txtgoto hasn't got a value then a message box is displayed
If txtGoto.Text = "" Then
    MsgBox "Enter a number to find", vbOKOnly
        Else
            
    If txtGoto.Text > intNumstore Then
        MsgBox "Max Value 843", vbOKOnly
            txtGoto.Text = ""
                
                Else
            
        'takes away 1 from the number inputted into the text box
        intNumgoto = txtGoto.Text - 1

        'displays it
        List1.ListIndex = intNumgoto
            
            txtGoto.Text = ""
            intNumgoto = 0
End If
End If

errorrepair:
error_Cancel

End Sub

Private Sub cmdPlay_Click()
Call List1_DblClick
End Sub

Private Sub cmdSearch_Click()
Unload Me
frmSearch.Show
End Sub

Private Sub Command1_Click()
ShellExecute 0, vbNullString, "https://github.com/kramwell/SegaClassics", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub Command2_Click()
Dim MyRandomNumber As Long 'The chosen number
Dim RandomMax As Long 'top end of range to pick from
Dim RandomMin As Long 'low end of range to pick from

RandomMin = 1
RandomMax = 843

Randomize Timer
MyRandomNumber = Int(Rnd(1) * RandomMax) + RandomMin
List1.ListIndex = MyRandomNumber - 1

End Sub

Private Sub Form_Load()
If App.PrevInstance Then
    MsgBox "Application already running", vbOKOnly
    End
End If
End Sub

Private Sub List1_Click()
cmdPlay.Default = True

'gets number from game selected and adds 1 to it
lblGameno.Caption = List1.ListIndex + 1

End Sub

Private Sub List1_DblClick()
Dim taskid As Long
Dim strPathname As String
Dim strFullname As String

On Error GoTo errorrepair


'gets the path name selected
strPathname = List1.Text
    
If strPathname = "" Then
    MsgBox "Select a game to play!", vbOKOnly
    Exit Sub
        Else
        
'makes strfullname hold the full path to the game
strFullname = """SMD\" & strPathname & " # SMD.ZIP"""
    
'opens the game you selected
taskid = Shell("32\Fusion\Fusion.exe " & strFullname, vbNormalFocus)

End If

errorrepair:
error_Cancel

End Sub

Private Sub error_Cancel()
If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "Error " & Err.Number
End If
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
'makes the command button the default on the text box
cmdGoto.Default = True
End Sub
