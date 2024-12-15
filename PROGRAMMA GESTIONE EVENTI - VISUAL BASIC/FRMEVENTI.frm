VERSION 5.00
Begin VB.Form FRMEVENTI 
   BorderStyle     =   0  'None
   Caption         =   "EVENTO"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleMode       =   0  'User
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CBOM 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4800
      TabIndex        =   24
      Text            =   "MIN"
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox CBOH 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FRMEVENTI.frx":0000
      Left            =   4080
      List            =   "FRMEVENTI.frx":0002
      TabIndex        =   23
      Text            =   "HH"
      Top             =   2640
      Width           =   735
   End
   Begin VB.Frame FMESUPPORTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.TextBox TXTPERSONE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MaxLength       =   16
         TabIndex        =   41
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TXTNC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaxLength       =   16
         TabIndex        =   38
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox CBOLC 
         Height          =   315
         ItemData        =   "FRMEVENTI.frx":0004
         Left            =   3000
         List            =   "FRMEVENTI.frx":0006
         TabIndex        =   36
         Text            =   "SELEZIONA"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Frame FMEOTHER 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   1095
         Begin VB.CheckBox CHKDJ 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "DJ"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox CHKMUSICA 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "MUSICA"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox CHKALCOL 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "ALCOL"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox CHKCIBO 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "CIBO"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.TextBox TXTLN 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaxLength       =   16
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox CBOL 
         Height          =   315
         ItemData        =   "FRMEVENTI.frx":0008
         Left            =   3000
         List            =   "FRMEVENTI.frx":000A
         TabIndex        =   28
         Text            =   "SELEZIONA"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TXTNOME 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         MaxLength       =   16
         TabIndex        =   26
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox CBOLE 
         Height          =   315
         ItemData        =   "FRMEVENTI.frx":000C
         Left            =   3000
         List            =   "FRMEVENTI.frx":000E
         TabIndex        =   22
         Text            =   "SELEZIONA"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox CBOLL 
         Height          =   315
         ItemData        =   "FRMEVENTI.frx":0010
         Left            =   1440
         List            =   "FRMEVENTI.frx":0029
         TabIndex        =   20
         Text            =   "SELEZIONA"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox CBOTE 
         Height          =   315
         ItemData        =   "FRMEVENTI.frx":0062
         Left            =   1440
         List            =   "FRMEVENTI.frx":0072
         TabIndex        =   19
         Text            =   "SELEZIONA"
         Top             =   720
         Width           =   1335
      End
      Begin VB.PictureBox ListView1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   120
         MousePointer    =   1  'Arrow
         ScaleHeight     =   1545
         ScaleWidth      =   6705
         TabIndex        =   13
         Top             =   3000
         Width           =   6735
      End
      Begin VB.Frame FMETITOLO 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   7095
         Begin VB.CommandButton CMDBACK 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   255
         End
         Begin VB.Label LBLTITOLO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1725
            TabIndex        =   12
            Top             =   0
            Width           =   90
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "EVENTI"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   240
            TabIndex        =   11
            Top             =   120
            Width           =   960
         End
      End
      Begin VB.CommandButton CMDNUOVO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NUOVO"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton CMDSALVA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SALVA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton CMDMODIFICA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "MODIFICA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton CMDELIMINA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ELIMINA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox TXTN 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   16
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   0
         Top             =   4920
      End
      Begin VB.ComboBox CBOGGC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "FRMEVENTI.frx":009E
         Left            =   1440
         List            =   "FRMEVENTI.frx":00A0
         TabIndex        =   3
         Text            =   "GG"
         Top             =   2640
         Width           =   615
      End
      Begin VB.ComboBox CBOMMC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2040
         TabIndex        =   2
         Text            =   "MM"
         Top             =   2640
         Width           =   615
      End
      Begin VB.ComboBox CBOYYC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2640
         TabIndex        =   1
         Text            =   "YY"
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label LBLP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N. PERSONE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label LBLNC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   39
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label LBLC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LISTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label LBLLUOGON 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LISTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label LBLNOME 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   27
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LBLDATANASCITAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DATA E ORA DEL EVENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label LISTA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LISTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label LBLNE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N. EVENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LBLLUOGO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LUOGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label LBLIMMAGNE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "UPLOAD IMMAGINE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7680
         TabIndex        =   16
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label LBLTE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO EVENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATA MATRIMONIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   25
      Top             =   2160
      Width           =   2175
   End
End
Attribute VB_Name = "FRMEVENTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim RSM As Recordset
Dim RSA As Recordset
Dim RSP As Recordset
Dim RSC As Recordset
Dim A As Integer

Public Function CONTROLLACODICE(CODCOM As String) As Boolean
CERCACODICE = "N_E='" + CODCOM + "'"
RS.FindFirst (CERCACODICE)
If RS.NoMatch = True Then
CONTROLLACODICE = False
Else
CONTROLLACODICE = True
Exit Function
End If
End Function

Sub CARICADATA()
CBOYYC.AddItem Right(Date, 4)
For I = 1 To 60
    CBOYYC.AddItem Right(Date, 4) - I
    If I <= 31 Then
    CBOGGC.AddItem (I)
    End If
    If I <= 12 Then
        CBOMMC.AddItem (I)
    End If
Next I
End Sub

Sub CARICARECORD()
Dim LIST As ListItem
ListView1.ListItems.Clear
RS.MoveFirst
Do Until RS.EOF
Set LIST = ListView1.ListItems.Add(, , RS!N_E)
If RS!CIBO = True Then
    LIST.SubItems(1) = RS!CIBO
End If
If RS!ALCOL = True Then
    LIST.SubItems(2) = RS!ALCOL
End If
If RS!MUSICA = True Then
    LIST.SubItems(3) = RS!MUSICA
End If
If RS!DJ = True Then
    LIST.SubItems(4) = RS!DJ
End If
LIST.SubItems(5) = RS!Data
LIST.SubItems(6) = RS!ORA
LIST.SubItems(7) = RS!PERSONE
LIST.SubItems(8) = RS!TIPO
LIST.SubItems(9) = RS!LUOGO
LIST.SubItems(10) = RS!CANTANTE
RS.MoveNext
Loop
End Sub

Sub PULISCI()
TXTN.Text = ""
TXTNOME.Text = ""
TXTLN.Text = ""
TXTNC.Text = ""
CBOLE.Text = "SELEZIONA"
CBOL.Text = "SELEZIONA"
CBOLC.Text = "SELEZIONA"
CBOTE.Text = "SELEZIONA"
CBOLL.Text = "SELEZIONA"
TXTPERSONE.Text = ""
CHKCIBO.Value = 0
CHKMUSICA.Value = 0
CHKALCOL.Value = 0
CHKDJ.Value = 0
CBOGGC.Text = "GG"
CBOMMC.Text = "MM"
CBOYYC.Text = "YY"
CBOH.Text = "HH"
CBOM.Text = "MIN"
End Sub

Sub TRASFERISCI()
RS("N_E") = TXTN.Text
If CHKCIBO.Value = 1 Then
    RS("CIBO") = True
End If
If CHKMUSICA.Value = 1 Then
    RS("MUSICA") = True
End If
If CHKALCOL.Value = 1 Then
    RS("ALCOL") = True
End If
If CHKDJ.Value = 1 Then
    RS("DJ") = True
End If
If CHKCIBO.Value = 1 Then
    RS("CIBO") = True
End If
If CBOLE.Text <> "SELEZIONA" Then
    RS("TIPO") = CBOLE.Text
End If
If CBOL.Text <> "SELEZIONA" Then
    RS("LUOGO") = CBOL.Text
End If
If CBOLC.Text <> "SELEZIONA" Then
    RS("CANTANTE") = CBOLC.Text
End If
If CBOGGC.Text <> "GG" And CBOGGC.Text <> "MM" And CBOGGC.Text <> "YY" Then
    RS("DATA") = CBOGGC.Text + "/" + CBOMMC.Text + "/" + CBOYYC.Text
End If
If CBOH.Text <> "HH" And CBOM.Text <> "MIN" Then
    RS("ORA") = CBOGGC.Text + ":" + CBOMMC.Text
End If
RS("PERSONE") = TXTPERSONE.Text

End Sub


Private Sub CBOL_Change()
If CBOL.Text <> "SELEZIONA" Then
    SQLC = ("SELECT NOME_C " & " FROM T_CLIENTI " & " WHERE CF_C = '" & (CBOLC.Text) & "'")
    Set RSC = DB.OpenRecordset(SQLC)
    If RSC.RecordCount <> 0 Then
        TXTNC.Text = RSC("NOME_C")
        SQL = ""
    End If
    SQLC = ""
    RSC.Close
End If
End Sub

Private Sub CBOLC_Change()
If CBOLC.Text <> "SELEZIONA" Then
    SQLC = ("SELECT NOME_C " & " FROM T_CLIENTI " & " WHERE CF_C = '" & (CBOLC.Text) & "'")
    Set RSC = DB.OpenRecordset(SQLC)
    If RSC.RecordCount <> 0 Then
        TXTNC.Text = RSC("NOME_C")
        SQL = ""
    End If
    SQLC = ""
    RSC.Close
End If
End Sub




Private Sub TXTN_LostFocus()
If CONTROLLACODICE(TXTN) = True Then
MsgBox ("CODISE FISCALE CLIENTE ESISTENTE")
TXTN.Text = ""
TXTN.SetFocus
End If
End Sub


Private Sub CBOLE_Click()
CBOLE.Clear
If CBOTE.Text = "MATRIMONIO" Then
    CBOLE.Clear
    SQLM = ("SELECT N_M, T_M " & " FROM T_MATRIMONIO ")
    If RSM.RecordCount <> 0 Then
    SQLM = ("SELECT N_M, T_M " & " FROM T_MATRIMONIO ")
    Set RSM = DB.OpenRecordset(SQLM)
        RSM.MoveFirst
        While Not RSM.EOF
            If RSM("T_M") = CBOLE.Text Then
                TXTNOME.Text = RSM!T_M
            End If
            RSM.MoveNext
        Wend
    RSM.Close
    End If
ElseIf CBOTE.Text = "ANNIVERSARIO" Then
    CBOLE.Clear
    Set RSA = DB.OpenRecordset("T_ANNIVERSARIO", dbOpenDynaset)
    If RSA.RecordCount <> 0 Then
    SQLM = ("SELECT N_A, T_A " & " FROM T_ANNIVERSARIO ")
    Set RSA = DB.OpenRecordset(SQLM)
        RSA.MoveFirst
        While Not RSA.EOF
            If RSA("T_A") = CBOLE.Text Then
                TXTNOME.Text = RSA!T_A
            End If
            RSA.MoveNext
        Wend
    RSA.Close
    End If
ElseIf CBOTE.Text = "PARTY" Then
    CBOLE.Clear
    Set RSP = DB.OpenRecordset("T_PARTY", dbOpenDynaset)

    If RSP.RecordCount <> 0 Then
    SQLM = ("SELECT N_P, T_P " & " FROM T_PARTY ")
    Set RSP = DB.OpenRecordset(SQLM)
        RSP.MoveFirst
        While Not RSP.EOF
            If RSP!T_P = CBOLE.Text Then
                TXTNOME.Text = RSP!T_P
            End If
            RSP.MoveNext
        Wend
    RSP.Close
    End If
End If
End Sub

Private Sub CBOTE_Click()
CBOLE.Clear
If CBOTE.Text = "MATRIMONIO" Then
    CBOLE.Clear
    Set RSM = DB.OpenRecordset("T_MATRIMONIO", dbOpenDynaset)
    If RSM.RecordCount <> 0 Then
    SQLM = ("SELECT N_M " & " FROM T_MATRIMONIO ")
    Set RSM = DB.OpenRecordset(SQLM)
        RSM.MoveFirst
        While Not RSM.EOF
            CBOLE.AddItem RSC("N_M")
            RSM.MoveNext
        Wend
    RSM.Close
    End If
ElseIf CBOTE.Text = "ANNIVERSARIO" Then
    CBOLE.Clear
    Set RSA = DB.OpenRecordset("T_ANNIVERSARIO", dbOpenDynaset)
    If RSA.RecordCount <> 0 Then
    SQLM = ("SELECT N_A " & " FROM T_ANNIVERSARIO ")
    Set RSA = DB.OpenRecordset(SQLM)
        RSA.MoveFirst
        While Not RSA.EOF
            CBOLE.AddItem RSA("N_A")
            RSA.MoveNext
        Wend
    RSA.Close
    End If
ElseIf CBOTE.Text = "PARTY" Then
    CBOLE.Clear
    Set RSP = DB.OpenRecordset("T_PARTY", dbOpenDynaset)
    If RSP.RecordCount <> 0 Then
    SQLM = ("SELECT N_P " & " FROM T_PARTY ")
    Set RSP = DB.OpenRecordset(SQLM)
        RSP.MoveFirst
        While Not RSP.EOF
            CBOLE.AddItem RSP("N_P")
            RSP.MoveNext
        Wend
    RSP.Close
    End If
End If
End Sub

Private Sub CHKMUSICA_Click()
If CHKMUSICA.Value = 1 Then
    CHKDJ.Visible = True
    LBLC.Visible = True
    LBLNC.Visible = True
    CBOLC.Visible = True
    TXTNC.Visible = True
Else
    CHKDJ.Visible = False
    LBLC.Visible = False
    LBLNC.Visible = False
    CBOLC.Visible = False
    TXTNC.Visible = False
End If
    
End Sub

Private Sub CMDBACK_Click()
A = 1
FRMHOME.Show
End Sub
Sub CARICACANTANTI()
SQLC = ("SELECT CF_C " & " FROM T_CANTANTI ")
Set RSC = DB.OpenRecordset(SQLC)
    RSC.MoveFirst
    While Not RSC.EOF
        CBOLC.AddItem RSC("CF_C")
        RSC.MoveNext
    Wend
RSC.Close
End Sub

Private Sub Form_ACTIVATE()
Timer1.Enabled = True
FRMEVENTI.Top = 8000
FRMEVENTI.Height = 0
A = 0
CARICADATA
Set DB = Workspaces(0).OpenDatabase(App.Path + "\DATABASE")
Set RS = DB.OpenRecordset("T_EVENTI", dbOpenDynaset)
If RS.RecordCount <> 0 Then
    CARICARECORD
End If
CARICACANTANTI
End Sub

Private Sub CMDSALVA_Click()
If TXTN.Text Then
RS.AddNew
TRASFERISCI
RS.Update
PULISCI
CARICARECORD
CMDNUOVO.Enabled = True
CMDSALVA.Enabled = False
Else
MsgBox ("MODULO IN COMPLETO, IMPOSSIBILE SALVARE IL CLIENTE")
CMDNUOVO.Enabled = True
End If
End Sub

Private Sub CMDNUOVO_Click()
PULISCI
ListView1.Refresh
CMDSALVA.Enabled = True
CMDNUOVO.Enabled = False
CMDMODIFICA.Enabled = False
End Sub

Private Sub CMDMODIFICA_Click()
C = MsgBox("SEI SICURO DI VOLER MODIFICARE IL RECORD", vbYesNo)
If C = 6 Then
    ricercamodifica = "N_E'" & TXTN.Text & "'"
    RS.FindFirst (ricercamodifica)
    If RS.NoMatch = False Then
        RS.Edit
        TRASFERISCI
        RS.Update
        PULISCI
        CMDMODIFICA.Enabled = False
        CMDNUOVO.Enabled = True
        CARICARECORD
    End If
End If

End Sub

Private Sub CMDELIMINA_Click()
RISPOSTA = MsgBox("sei sicuro di voler eliminare il record", vbYesNo)
If RISPOSTA = 6 Then
    ricercacfc = "N_E='" & TXTN.Text & "'"
    RS.FindFirst (ricercacfc)
    If RS.NoMatch = False Then
    RS.Delete
    PULISCI
    CARICARECORD
    End If
End If
End Sub


Private Sub Timer1_Timer()
If A = 0 Then
    If FRMEVENTI.Height > 4400 Then
        FRMEVENTI.Height = 4890
    Else
        FRMEVENTI.Height = FRMEVENTI.Height + 200
        FRMEVENTI.Top = FRMEVENTI.Top - 200

    End If
ElseIf A = 1 Then
    If FRMEVENTI.Height < 100 Then
        Unload Me
    Else
        FRMEVENTI.Height = FRMEVENTI.Height - 200
        FRMEVENTI.Top = FRMEVENTI.Top + 200
    End If

End If
End Sub


