VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMRISTORANTE 
   BorderStyle     =   0  'None
   Caption         =   "RESTORAZIONE"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7110
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
      Begin VB.TextBox TXTSTATO 
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
         Left            =   3480
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Frame FMETITOLO 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   19
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
            TabIndex        =   20
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
            Caption         =   "RISTORANTE"
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
            Left            =   135
            TabIndex        =   21
            Top             =   120
            Width           =   1650
         End
      End
      Begin VB.TextBox TXTPIVA 
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
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TXTNCIVICO 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TXTVIA 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TXTPROVINCIA 
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
         Left            =   2280
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TXTCITTA 
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
         TabIndex        =   14
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TXTCAP 
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
         Left            =   6120
         TabIndex        =   13
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TXTPREFISSOTEL 
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
         MaxLength       =   4
         TabIndex        =   12
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox TXTPREFISSOCEL 
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
         MaxLength       =   3
         TabIndex        =   11
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox TXTEMAIL 
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
         Left            =   3000
         TabIndex        =   10
         Top             =   2280
         Width           =   2295
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
         Left            =   120
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox TXTCF 
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
         Left            =   2280
         MaxLength       =   16
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox TXTSUFFISSOTEL 
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
         Left            =   600
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TXTSUFFISSOCEL 
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
         Left            =   2040
         MaxLength       =   7
         TabIndex        =   3
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TXTDEN 
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
         Left            =   4560
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   0
         Top             =   4920
      End
      Begin VB.TextBox TXTMT 
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
         Left            =   5400
         MaxLength       =   3
         TabIndex        =   1
         Top             =   2280
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   37
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
      Begin VB.Label LBLSTATO 
         BackStyle       =   0  'Transparent
         Caption         =   "STATO"
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
         Left            =   3480
         TabIndex        =   36
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label LBLCFC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CODICE FISCALE"
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
         Left            =   2280
         TabIndex        =   35
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label LBLNOMEC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PARTITA IVA"
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
         TabIndex        =   34
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label LBLVIAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "VIA / PIAZZA"
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
         TabIndex        =   33
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LBLNCIVICOC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "N"
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
         Left            =   1680
         TabIndex        =   32
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label LBLCITTAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CITTA"
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
         TabIndex        =   31
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label LBLPROVINCIAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA"
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
         Left            =   2280
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LBLCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CAP"
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
         Left            =   6120
         TabIndex        =   29
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label LBLTELEFONOC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONO"
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
         TabIndex        =   28
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label LBLCELLULARE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CELLULARE"
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
         TabIndex        =   27
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label LBLEMAILC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL"
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
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
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
         TabIndex        =   25
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label LBLDEN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DENOMINAZIONE"
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
         Left            =   4560
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "METRI Q"
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
         Left            =   5400
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Timer Timer2 
      Left            =   720
      Top             =   1320
   End
End
Attribute VB_Name = "FRMRISTORANTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim A As Integer

Public Function CONTROLLACODICE(CODCOM As String) As Boolean
CERCACODICE = "PIVA='" + CODCOM + "'"
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
Set LIST = ListView1.ListItems.Add(, , RS!PIVA)
LIST.SubItems(1) = RS!CF
LIST.SubItems(2) = RS!DEN
LIST.SubItems(3) = RS!VIA
LIST.SubItems(4) = RS!N_CIVICO
LIST.SubItems(5) = RS!STATO
LIST.SubItems(6) = RS!CITTA
LIST.SubItems(7) = RS!CAP
LIST.SubItems(8) = RS!PREFISSOTEL
LIST.SubItems(9) = RS!SUFFISSOTEL
LIST.SubItems(10) = RS!PREFISSOCEL
LIST.SubItems(11) = RS!SUFFISSOCEL
LIST.SubItems(12) = RS!EMAIL
LIST.SubItems(13) = RS!METRI
RS.MoveNext
Loop
End Sub
Sub PULISCI()
TXTPIVA.Text = ""
TXTCF.Text = ""
TXTDEN.Text = ""
TXTVIA.Text = ""
TXTNCIVICO.Text = ""
TXTCITTA.Text = ""
TXTPROVINCIA.Text = ""
TXTCAP.Text = ""
TXTSTATO.Text = ""
TXTPREFISSOTEL.Text = ""
TXTSUFFISSOTEL.Text = ""
TXTPREFISSOCE.Text = ""
TXTSUFFISSOCE.Text = ""
TXTEMAIL.Text = ""
TXTMT.Text = ""
End Sub
Sub TRASFERISCI()
RS("CF_C") = TXTPIVA.Text
RS("NOME_C") = TCTCF.Text
RS("COGNOME_C") = TXTDEN.Text
RS("VIA_C") = TXTVIA.Text
RS("N_CIVICO_C") = TXTNCIVICO.Text
RS("CITTA_C") = TXTCITTA.Text
RS("PROVINCIA_C") = TXTPROVINCIA.Text
RS("CAP_C") = Val(TXTCAP.Text)
If TXTPREFISSOTEL.Text = "" Then
RS("PREFISSOTEL") = 0
Else
RS("PREFISSOTEL") = TXTPREFISSOTEL.Text
End If
RS("SUFFISSOTEL") = TXTSUFFISSOTEL.Text
RS("PREFISSOCEL_C") = TXTPREFISSOCEL.Text
RS("SUFFISSOC_C") = TXTSUFFISSOCELC.Text
RS("EMAIL") = TXTEMAIL.Text
RS("METRI") = TXTMT.Text
End Sub



Private Sub TXTCFC_LostFocus()
If CONTROLLACODICE(TXTPIVA) = True Then
MsgBox ("CODISE FISCALE CLIENTE ESISTENTE")
TXTPIVA.Text = ""
TXTPIVA.SetFocus
End If
End Sub

Private Sub CMDBACK_Click()
A = 1
FRMHOME.Show
End Sub


Private Sub Form_ACTIVATE()
Timer1.Enabled = True
FRMRISTORANTE.Top = 8000
FRMRISTORANTE.Height = 0
A = 0

Set DB = Workspaces(0).OpenDatabase(App.Path + "\DATABASE")
Set RS = DB.OpenRecordset("T_RISTORANTE", dbOpenDynaset)
If RS.RecordCount <> 0 Then
CARICARECORD
End If
End Sub

Private Sub CMDSALVA_Click()
If TXTPIVA.Text And TXTCF.Text <> "" And TCTDEN.Text <> "" And TXTVIA.Text <> "" And TXTCAP.Text <> "" And TXTCITTA.Text <> "" And TXTPROVINCIA.Text <> "" And TXTSTATO.Text <> "" And TXTPREFISSOTEL.Text <> "" And TXTSUFFISSOTEL.Text <> "" And TXTPREFISSOCEL.Text <> "" And TXTSUFFISSOCEL.Text <> "" Then
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
    ricercamodifica = "PIVA='" & TXTPIVA.Text & "'"
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
    ricercacfc = "PIVA='" & TXTPIVA.Text & "'"
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
    If FRMRISTORANTE.Height > 4400 Then
        FRMRISTORANTE.Height = 4890
    Else
        FRMRISTORANTE.Height = FRMRISTORANTE.Height + 200
        FRMRISTORANTE.Top = FRMRISTORANTE.Top - 200

    End If
ElseIf A = 1 Then
    If FRMRISTORANTE.Height < 100 Then
        Unload Me
    Else
        FRMRISTORANTE.Height = FRMRISTORANTE.Height - 200
        FRMRISTORANTE.Top = FRMRISTORANTE.Top + 200
    End If

End If
End Sub


