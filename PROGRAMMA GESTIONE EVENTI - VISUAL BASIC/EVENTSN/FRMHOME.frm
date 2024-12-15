VERSION 5.00
Begin VB.Form FRMHOME 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "HOME"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   -240
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   16920
      Top             =   9600
   End
   Begin VB.Frame FMESUPPORTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6735
      Begin VB.CommandButton CMDEVENTI 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDANNIVERSARIO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDFATTURETAB 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "FT"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDFATTURA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "FD"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDMATRIMONIO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDPARTY 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4200
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDDISCOTECA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3120
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDCITTA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDBAR 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDCANTANTI 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDRISTORANTE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2760
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDHOTEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5280
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDFORNITORI 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1920
         MaskColor       =   &H00FFFF80&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton CMDCLIENTE 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         MaskColor       =   &H00FFFF80&
         Picture         =   "FRMHOME.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.Shape SHVISURA 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   735
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton CMDEXIT 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Label LBLNOME 
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label LBLTITOLO2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO ORG"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   960
      TabIndex        =   6
      Top             =   0
      Width           =   4920
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   3720
      TabIndex        =   5
      Top             =   5760
      Width           =   2040
   End
   Begin VB.Shape SHPUPLINE 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label LBLDATAORA 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   13800
      TabIndex        =   4
      Top             =   9360
      Width           =   2925
   End
End
Attribute VB_Name = "FRMHOME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer

Private Sub CMDFATTURETAB_Click()
A = 1
FRMFATTURETAB.Show
End Sub

Private Sub CMDFORNITORI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDFORNITORI.Top
SHVISURA.Left = CMDFORNITORI.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDCLIENTE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDCLIENTE.Top
SHVISURA.Left = CMDCLIENTE.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDFORNITORI_Click()
A = 1
FRMFORNITORI.Show
End Sub
Private Sub CMDANNIVERSARIO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDANNIVERSARIO.Top
SHVISURA.Left = CMDANNIVERSARIO.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDANNIVERSARIO_Click()
A = 1
FRMANNIVERSARIO.Show
End Sub
Private Sub CMDCLIENTE_Click()
A = 1
FRMCLIENTI.Show
End Sub
Private Sub CMDEVENTI_Click()
A = 1
FRMEVENTI.Show
End Sub
Private Sub CMDEVENTI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDEVENTI.Top
SHVISURA.Left = CMDEVENTI.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDHOTEL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDHOTEL.Top
SHVISURA.Left = CMDHOTEL.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDHOTEL_Click()
A = 1
FRMHOTEL.Show
End Sub
Private Sub CMDBAR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDBAR.Top
SHVISURA.Left = CMDBAR.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDBAR_Click()
A = 1
FRMBAR.Show
End Sub
Private Sub CMDEXIT_Click()
End
End Sub
Private Sub CMDFATTURA_Click()
A = 1
FRMFATTURA.Show
End Sub
Private Sub CMDFATTURA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDFATTURA.Top
SHVISURA.Left = CMDFATTURA.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDCITTA_Click()
A = 1
FRMCITTA.Show
End Sub
Private Sub CMDCITTA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDCITTA.Top
SHVISURA.Left = CMDCITTA.Left - 135
SHVISURA.Visible = True
End Sub

Private Sub CMDDISCOTECA_Click()
A = 1
FRMDISCOTECA.Show
End Sub

Private Sub CMDDISCOTECA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDDISCOTECA.Top
SHVISURA.Left = CMDDISCOTECA.Left - 135
SHVISURA.Visible = True
End Sub

Private Sub CMDPARTY_Click()
A = 1
FRMPARTY.Show
End Sub

Private Sub CMDPARTY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDPARTY.Top
SHVISURA.Left = CMDPARTY.Left - 135
SHVISURA.Visible = True
End Sub

Private Sub CMDMATRIMONIO_Click()
A = 1
FRMMATRIMONIO.Show
End Sub

Private Sub CMDMATRIMONIO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDMATRIMONIO.Top
SHVISURA.Left = CMDMATRIMONIO.Left - 135
SHVISURA.Visible = True
End Sub
Private Sub CMDCANTANTI_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDCANTANTI.Top
SHVISURA.Left = CMDCANTANTI.Left - 135
SHVISURA.Visible = True
End Sub

Private Sub CMDCANTANTI_Click()
A = 1
FRMCANTANTE.Show
End Sub

Private Sub CMDRESTORANTE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDRESTORANTE.Top
SHVISURA.Left = CMDRESTORANTE.Left - 135
SHVISURA.Visible = True
End Sub

Private Sub CMDRISTORANTE_Click()
A = 1
FRMRISTORANTE.Show
End Sub
Private Sub CMDPREESPRATICO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDPREESPRATICO.Top
SHVISURA.Left = CMDPREESPRATICO.Left - 135
SHVISURA.Visible = True
End Sub



Private Sub CMDSIDEBAR_Click()
If FMEMENU.Width = 3255 Then
    FMEMENU.Width = 375
    CMDSIDEBAR.Left = 480
    FMESUPPORTO.Left = 480
    FMESUPPORTO.Width = 16333
Else
    FMEMENU.Width = 3255
    CMDSIDEBAR.Left = 3360
    FMESUPPORTO.Left = 3360
    FMESUPPORTO.Width = 13455
    
End If
End Sub

Private Sub FMESUPPORTO_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDCLIENTE.Top
SHVISURA.Left = CMDCLIENTE.Left - 135
SHVISURA.Visible = False
End Sub

Private Sub Form_ACTIVATE()
Timer1.Enabled = True
FRMHOME.Height = 0
A = 0
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SHVISURA.Top = CMDCLIENTE.Top
SHVISURA.Left = CMDCLIENTE.Left - 135
SHVISURA.Visible = False
End Sub
Private Sub Timer1_Timer()
LBLDATAORA.Caption = "ORE: " & Time & " DATA: " & Date
If A = 0 Then
    If FRMHOME.Height > 4890 Then
    A = 3
    Else
        FRMHOME.Height = FRMHOME.Height + 200
    End If
ElseIf A = 1 Then
    If FRMHOME.Height > 20 Then
        FRMHOME.Height = FRMHOME.Height - 200
    Else
        Unload Me
    End If

End If
End Sub




