VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paul Pickard's I-Gear & LAN Internet Unblocking System"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   11370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer5 
      Interval        =   4000
      Left            =   1920
      Top             =   1440
   End
   Begin VB.Timer Timer4 
      Interval        =   4000
      Left            =   1920
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   960
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1200
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   6360
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H009AC635&
      Caption         =   "Instructions"
      Height          =   255
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   480
      Width           =   10335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H009AC635&
      Caption         =   "Load Reconstructed Page"
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H009AC635&
      Caption         =   "Reconstruct Page"
      Height          =   255
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   11175
      ExtentX         =   19711
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009AC635&
      Caption         =   "Get/Decrypt Code"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox URL 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "http://www.yahoo.com"
      Top             =   120
      Width           =   3975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait While Paul Pickard's I-Gear And LAN Internet Unblocking System Collects And Syncronises Needed Data..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   11175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Decrypted Code"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Programmed By Paul Pickard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   8280
      TabIndex        =   9
      Top             =   7440
      Width           =   3120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please Wait While LAN Internet Unblocking System Collects And Syncronises Needed Data..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   -120
      TabIndex        =   8
      Top             =   7440
      Width           =   8520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label2.Caption = "Getting Webpage Coding, Decrypting Webpage Code..."
    On Error Resume Next
    
    Dim txt As String
    Dim b() As Byte
    
    Command1.Enabled = False
    
    b() = Inet1.OpenURL(URL.Text, 1)
    
    txt = ""
    


    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next

    Text2 = txt
    Command1.Enabled = True
    Label2.Caption = "Webpage Coding Decrypted, Click Reconstruct Page To Continue..."
    Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo asd
Open (App.Path & "\decrypt.html") For Output As #1
Print #1, Text2
Close #1
GoTo dsa
asd:
MsgBox "Oh My Life! An Error Occurred!"
dsa:
Label2.Caption = "Reconstruction Complete, Click Load Reconstructed Page To Continue..."
End Sub

Private Sub Command3_Click()
WebBrowser1.Navigate (App.Path & "\decrypt.html")
Label2.Caption = "Your Unblocked Page Has Been Loaded, Happy Surfing! - Paul Pickard"
End Sub

Private Sub Command4_Click()
MsgBox "Instructions:" & vbNewLine & "(A) Type In The Web Address You Wish To Access Which Is Blocked" & vbNewLine & "(B) Click Get/Decrypt Code" & vbNewLine & "(C) Click Reconstruct Page" & vbNewLine & "(D) Click Load Reconstructed Page, Then Surf Away!", vbInformation, "Instructions"
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Remember!, If This Program Worked For You And You Like It Please Give Credit To Paul Pickard As He Made This Program For Your Benifit - Thanks Paul Pickard", vbInformation, "Credit"
End Sub

Private Sub Label2_Change()
Timer3.Enabled = True
End Sub

Private Sub Timer2_Timer()
Text1.Text = Text1 + 1
End Sub

Private Sub Timer1_Timer()
If Text1.Text = "6" Then
Label2.Caption = "LAN Internet Unblocking System Collected Needed Data, Waiting For Input..."
Timer1.Enabled = False
Else
End If
End Sub

Private Sub Timer3_Timer()
Timer4.Enabled = True
If Label2.BackColor = &HC00000 Then
Label2.BackColor = &HFFFF&
Label2.ForeColor = &HFF&
Else
Label2.BackColor = &HC00000
Label2.ForeColor = &HFFFF00
End If
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
Timer3.Enabled = False
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
WebBrowser1.Visible = True
WebBrowser1.Navigate "http://www.pts.fire-bug.co.uk/unblock.htm"
Timer5.Enabled = False
End Sub
