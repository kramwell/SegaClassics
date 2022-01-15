VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Sega Game Search"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   Icon            =   "frmMore.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Play"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Back"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   2790
      ItemData        =   "frmMore.frx":000C
      Left            =   120
      List            =   "frmMore.frx":000E
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search for a Game"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Go"
         Height          =   255
         Left            =   2280
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBack_Click()
Unload Me
SegaClassics.Show
End Sub

Private Sub cmdPlay_Click()
Call lstSearch_DblClick
End Sub

Private Sub cmdSearch_Click()
Dim strLstname As String
Dim intCount As Integer
Dim intNumchar As Integer
Dim strSearch As String
Dim intSearchchar As Integer
Dim strNowlstname As String
Dim intStartpoint As Integer
Dim intEndpoint As Integer
Dim intUntil As Integer
Dim strSearchbf As String

lstSearch.Clear

Do

'selects the first line in the list
SegaClassics.List1.Selected(intCount) = True

'adds 1 so next time it will select the 2nd line
intCount = intCount + 1

'makes the line name take the value of a varible
strLstname = SegaClassics.List1.Text
strSearchbf = strLstname
strLstname = LCase(strLstname)

'gets the number of charactors in the sentance
intNumchar = Len(strLstname)

'gets the name of the search query imputted and puts it into a varible
strSearch = LCase(txtSearch.Text)

'gets the number of charactors in the search
intSearchchar = Len(strSearch)

'makes the number of charactors in the search query equal to intendpoint
intEndpoint = intSearchchar
intUntil = intSearchchar

    '-=-=-=
    'Loop this inside a loop
Do


'this makes intstartpoint increase 1
 intStartpoint = intStartpoint + 1


'this gets the list name, the first number of intstartpoint and the first number lengh
'of the value inputed then displays the word in between the lengh and displays it
strNowlstname = Mid(strLstname, intStartpoint, intEndpoint)



If strNowlstname = strSearch Then

'if no value is entered in the text box then a message box appears and
'ends the run
If strNowlstname = "" Then
    MsgBox "Enter a word to search", vbOKOnly, "Search Error"
    Exit Sub
        Else
    
    lstSearch.AddItem strSearchbf
    

    
    intUntil = intNumchar

End If
End If

intUntil = intUntil + 1
    'end the loop
    '-=-=-=
'intendpoint = intendpoint + 1
Loop Until intUntil > intNumchar

'resets the start point to its original value
intStartpoint = 0

Loop Until intCount = 843

If intCount = 843 Then
Unload SegaClassics
End If

If lstSearch.ListCount = 0 Then
MsgBox "The search returned NO results", vbOKOnly, "Search Results"
txtSearch.Text = ""
End If

End Sub

Private Sub Form_Load()
If App.PrevInstance Then
    MsgBox "Application already running", vbOKOnly
    End
End If
End Sub

Private Sub lstSearch_Click()
cmdPlay.Default = True
End Sub

Private Sub lstSearch_DblClick()
Dim taskid As Long
Dim strPathname As String
Dim strFullname As String

On Error GoTo errorrepair

'gets the path name selected
strPathname = lstSearch.Text
    
If strPathname = "" Then
    MsgBox "Select a game to play!", vbOKOnly, "Play Error"
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

Private Sub txtSearch_Click()
cmdSearch.Default = True
End Sub
