VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMFATTURETAB 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FMESUPPORTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton CMDBACK 
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
         Height          =   375
         Left            =   6960
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TXTCFATT 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox TXTIMPONIBILE 
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox TXTIVA 
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox TXTIMPOSTA 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox TXTPFINALE 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   2760
         Width           =   2415
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5040
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5040
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   5640
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   360
         TabIndex        =   11
         Top             =   3240
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PIVA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "DEN"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "VIA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "STATO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CITTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CAP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "PRE TEL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "SUF TEL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "PRE CEL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "SUF CEL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "EMAIL"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "METRI"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label TITOLO 
         BackColor       =   &H00000000&
         Caption         =   "FATTURE TABELLARI"
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
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Inserisci i dati della fattura tabellare:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "CODICE FATTURA"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "IMPONIBILE"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "ALIQUOTA IVA"
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
         Height          =   375
         Left            =   5520
         TabIndex        =   14
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "IMPORTO IMPOSTA"
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
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "PREZZO FINALE"
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
         Left            =   2880
         TabIndex        =   12
         Top             =   2400
         Width           =   2415
      End
   End
End
Attribute VB_Name = "FRMFATTURETAB"
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
CERCACODICE = "CODICE_FATTURAT='" + CODCOM + "'"
RS.FindFirst (CERCACODICE)
If RS.NoMatch = True Then
CONTROLLACODICE = False
Else
CONTROLLACODICE = True
Exit Function
End If
End Function
Sub CARICARECORD()
Dim LIST As ListItem
ListView1.ListItems.Clear
RS.MoveFirst
Do Until RS.EOF
Set LIST = ListView1.ListItems.Add(, , RS!CODICE_FATTURAT)
LIST.SubItems(1) = RS!IMPONIBILE_FATTURAT
LIST.SubItems(2) = RS!ALIQUOTA
LIST.SubItems(3) = RS!IMPOSTA
LIST.SubItems(4) = RS!PREZZO_FINALE
RS.MoveNext
Loop
End Sub
Sub PULISCI()
TXTCFATT.Text = ""
TXTIMPONIBILE.Text = ""
TXTIVA.Text = ""
TXTIMPOSTA.Text = ""
TXTPFINALE.Text = ""
End Sub
Sub TRASFERISCI()
RS("CODICE_FATTURAT") = TXTCFATT.Text
RS("IMPONIBILE_FATTURAT") = TXTIMPONIBILE.Text
RS("ALIQUOTA") = TXTIVA.Text
RS("IMPOSTA") = TXTIMPOSTA.Text
RS("VIA_C") = TXTVIAC.Text
RS("PREZZO_FINALE") = TXTPFINALE.Text
End Sub
Private Sub TXTCFATT_LostFocus()
If CONTROLLACODICE(TXTCFATT) = True Then
MsgBox ("CODISE FATTURA ESISTENTE")
TXTCFATT.Text = ""
TXTCFATT.SetFocus
End If
End Sub
Private Sub Form_activate()
Timer1.Enabled = True
FRMFATTURETAB.Top = 8000
FRMFATTURETAB.Height = 0
A = 0
Set DB = Workspaces(0).OpenDatabase(App.Path + "\DATABASE")
Set RS = DB.OpenRecordset("T_FATTAB", dbOpenDynaset)
If RS.RecordCount <> 0 Then
CARICARECORD
End If
End Sub

Private Sub CMDSALVA_Click()
If TXTCFATT.Text <> "" And TXTIMPONIBILE.Text <> "" And TXTIVA.Text <> "" And TXTIMPOSTA.Text <> "" And TXTPFINALE.Text <> "" Then
RS.AddNew
TRASFERISCI
RS.Update
PULISCI
CARICARECORD
CMDNUOVO.Enabled = True
CMDSALVA.Enabled = False
Else
MsgBox ("MODULO INCOMPLETO, IMPOSSIBILE SALVARE IL CLIENTE")
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
    ricercamodifica = "CODICE_FATTURAT='" & TXTCFATT.Text & "'"
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
    ricercacf = "CODICE_FATTURAT='" & TXTCFATT.Text & "'"
    RS.FindFirst (ricercacf)
    If RS.NoMatch = False Then
    RS.Delete
    PULISCI
    CARICARECORD
    End If
End If

End Sub

Private Sub Timer1_Timer()
If A = 0 Then
    If FRMFATTURETAB.Height > 4400 Then
        FRMFATTURETAB.Height = 4890
    Else
        FRMFATTURETAB.Height = FRMFATTURETAB.Height + 200
        FRMFATTURETAB.Top = FRMFATTURETAB.Top - 200

    End If
ElseIf A = 1 Then
    If FRMFATTURETAB.Height < 100 Then
        Unload Me
    Else
        FRMFATTURETAB.Height = FRMFATTURETAB.Height - 200
        FRMFATTURETAB.Top = FRMFATTURETAB.Top + 200
    End If
End If
End Sub


