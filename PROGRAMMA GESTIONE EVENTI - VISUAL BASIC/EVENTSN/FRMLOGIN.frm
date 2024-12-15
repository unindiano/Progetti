VERSION 5.00
Begin VB.Form FRMLOGIN 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6603.932
   ScaleMode       =   0  'User
   ScaleWidth      =   6137.679
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   3120
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -480
      TabIndex        =   5
      Top             =   0
      Width           =   5655
      Begin VB.Label LBLTITOLO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   480
         TabIndex        =   6
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.TextBox TXTPASSWORD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox TXTUSERNAME 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton CMDLOGIN 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton CMDEXIT 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.CommandButton CMDREGISTRATI 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "REGISTRATI"
      Height          =   255
      Left            =   3840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Image IMGUSER 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      Picture         =   "FRMLOGIN.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      Picture         =   "FRMLOGIN.frx":3319
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "FRMLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Private Sub cmdentrata_Click()
FRMLOGIN.Hide
FRMHOME.Show
End Sub

Private Sub CMDEXIT_Click()
End
End Sub

Private Sub CMDEXIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDEXIT.BackColor = &HC0FFC0
CMDLOGIN.BackColor = &HFF&
End Sub

Private Sub CMDLOGIN_Click()
    If TXTUSERNAME.Text = "" Or TXTPASSWORD.Text = "" Then
        MsgBox "INSERISCI I DATI RICHIESTI!", vbCritical, "PROCESSO ERRATO"
        Exit Sub
    ElseIf TXTUSERNAME.Text <> "" And TXTPASSWORD.Text <> "" Then
        B = " USERNAME_LOGIN='" + TXTUSERNAME.Text + "'"
        RS.FindFirst (B)
        If RS.NoMatch = False Then
            If RS("PASSWORD_LOGIN") = TXTPASSWORD.Text Then
               FRMHOME.LBLNOME.Caption = TXTUSERNAME.Text
                FRMHOME.Show
                Unload Me
            Else
            MsgBox ("DATI NON VALIDI")
            End If
        Else
        MsgBox ("DATI NON VALIDI")
        End If
                
    End If
End Sub

Private Sub CMDLOGIN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDLOGIN.BackColor = &HC0FFC0
CMDEXIT.BackColor = &HFF&
End Sub
Private Sub CMDREGISTRATI_Click()
If TXTUSERNAME.Text <> "" And TXTPASSWORD.Text <> "" Then
    RS!USERNAME_LOGIN = TXTUSERNAME.Text
    RS!PASSWORD_LOGIN = TXTPASSWORD.Text
    RS.Update
    TXTUSERNAME.Text = ""
    TXTPASSWORD.Text = ""
Else
    MsgBox ("INSERIRE I DATI !!!")
End If
End Sub
Private Sub Form_ACTIVATE()
Timer1.Enabled = True
Set DB = Workspaces(0).OpenDatabase(App.Path + "\DATABASE")
Set RS = DB.OpenRecordset("T_LOGIN", dbOpenDynaset)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CMDLOGIN.BackColor = &HFF&
CMDEXIT.BackColor = &HFF&
End Sub

Private Sub Timer1_Timer()
If LBLTITOLO1.ForeColor = vbWhite Then
    LBLTITOLO1.ForeColor = vbGreen
    LBLTITOLO2.ForeColor = vbWhite
ElseIf LBLTITOLO1.ForeColor = vbGreen Then
    LBLTITOLO1.ForeColor = vbYellow
    LBLTITOLO2.ForeColor = vbGreen
ElseIf LBLTITOLO1.ForeColor = vbYellow Then
    LBLTITOLO1.ForeColor = vbWhite
    LBLTITOLO2.ForeColor = vbYellow
End If
End Sub



