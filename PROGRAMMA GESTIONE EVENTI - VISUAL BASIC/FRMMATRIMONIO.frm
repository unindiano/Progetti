VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMMATRIMONIO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "MATRIMONIO"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7738.542
   ScaleMode       =   0  'User
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
         Left            =   1320
         TabIndex        =   19
         Text            =   "YY"
         Top             =   1800
         Width           =   975
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
         Left            =   720
         TabIndex        =   18
         Text            =   "MM"
         Top             =   1800
         Width           =   615
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
         ItemData        =   "FRMMATRIMONIO.frx":0000
         Left            =   120
         List            =   "FRMMATRIMONIO.frx":0002
         TabIndex        =   17
         Text            =   "GG"
         Top             =   1800
         Width           =   615
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   0
         Top             =   4920
      End
      Begin VB.TextBox TXTTM 
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
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   2295
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
         TabIndex        =   11
         Top             =   960
         Width           =   2175
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   4560
         Width           =   1215
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
         TabIndex        =   7
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox TXTC 
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
         Left            =   4800
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.Frame FMETITOLO 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "MATRIMONIO"
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
            Height          =   585
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1770
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
            TabIndex        =   4
            Top             =   0
            Width           =   90
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "REGISTRAZIONE"
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
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   1695
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   2990
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   255
         BackColor       =   -2147483643
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N_M"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TIPO_M"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "COMUNITA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DATA C."
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label LBLDATANASCITAC 
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
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label LBLDEN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO MATRIMONIO"
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
         TabIndex        =   16
         Top             =   720
         Width           =   2295
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
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label LBLEMAILC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "COMUNITA'"
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
         Left            =   4800
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label LBLCFC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO M."
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
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "FRMMATRIMONIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim A As Integer



Private Sub CMDBACK_Click()
A = 1
FRMHOME.Show
End Sub

Public Function CONTROLLACODICE(CODCOM As String) As Boolean
CERCACODICE = "N_M='" + CODCOM + "'"
RS.FindFirst (CERCACODICE)
If RS.NoMatch = True Then
CONTROLLACODICE = False
Else
CONTROLLACODICE = True
Exit Function
End If
End Function

Private Sub TXTCFC_LostFocus()
If CONTROLLACODICE(TXTN) = True Then
MsgBox ("CODISE FISCALE CLIENTE ESISTENTE")
TXTN.Text = ""
TXTN.SetFocus
End If
End Sub

Private Sub CMDSALVA_Click()
If TXTN.Text <> "" And TXTTM.Text <> "" And TXTC.Text <> "" Then
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

Sub PULISCI()
TXTN.Text = ""
TXTTM.Text = ""
TXTC.Text = ""
CBOGGC.Text = "GG"
CBOMMC.Text = "MM"
CBOYYC.Text = "YY"
End Sub

Sub TRASFERISCI()
RS("N_M") = TXTN.Text
RS("T_M") = TXTTM.Text
RS("C_M") = TXTC.Text
RS("DATA_M") = CBOGGC.Text + "/" + CBOMMC.Text + "/" + CBOYYC.Text
End Sub

Private Sub CMDMODIFICA_Click()
C = MsgBox("SEI SICURO DI VOLER MODIFICARE IL RECORD", vbYesNo)
If C = 6 Then
    ricercamodifica = "N_M='" & TXTN.Text & "'"
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
    ricercacfc = "N_M='" & TXTN.Text & "'"
    RS.FindFirst (ricercacfc)
    If RS.NoMatch = False Then
    RS.Delete
    PULISCI
    CARICARECORD
    End If
End If

End Sub


Private Sub Form_ACTIVATE()
Timer1.Enabled = True
FRMMATRIMONIO.Top = 8000
FRMMATRIMONIO.Height = 0
A = 0
CARICADATA
Set DB = Workspaces(0).OpenDatabase(App.Path + "\DATABASE")
Set RS = DB.OpenRecordset("T_MATRIMONIO", dbOpenDynaset)
If RS.RecordCount <> 0 Then
CARICARECORD
End If
End Sub

Sub CARICARECORD()
Dim LIST As ListItem
ListView1.ListItems.Clear
RS.MoveFirst
Do Until RS.EOF
Set LIST = ListView1.ListItems.Add(, , RS!N_A)
LIST.SubItems(1) = RS!T_A
LIST.SubItems(2) = RS!C_A
LIST.SubItems(3) = RS!DATA_A
RS.MoveNext
Loop
End Sub
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
Private Sub Timer1_Timer()
If A = 0 Then
    If FRMMATRIMONIO.Height > 4400 Then
        FRMMATRIMONIO.Height = 4890
    Else
        FRMMATRIMONIO.Height = FRMMATRIMONIO.Height + 200
        FRMMATRIMONIO.Top = FRMMATRIMONIO.Top - 200

    End If
ElseIf A = 1 Then
    If FRMMATRIMONIO.Height < 100 Then
        Unload Me
    Else
        FRMMATRIMONIO.Height = FRMMATRIMONIO.Height - 200
        FRMMATRIMONIO.Top = FRMMATRIMONIO.Top + 200
    End If

End If
End Sub


