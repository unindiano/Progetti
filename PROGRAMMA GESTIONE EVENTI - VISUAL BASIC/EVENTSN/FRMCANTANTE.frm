VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCANTANTE 
   BorderStyle     =   0  'None
   Caption         =   "CANTANTE"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
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
      Begin VB.TextBox TXTSUFFISSOCELC 
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
         Left            =   3600
         MaxLength       =   7
         TabIndex        =   27
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TXTSUFFISSOTC 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   26
         Top             =   2280
         Width           =   735
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
         Left            =   6000
         TabIndex        =   25
         Text            =   "YY"
         Top             =   840
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
         Left            =   5400
         TabIndex        =   24
         Text            =   "MM"
         Top             =   840
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
         ItemData        =   "FRMCANTANTE.frx":0000
         Left            =   4800
         List            =   "FRMCANTANTE.frx":0002
         TabIndex        =   23
         Text            =   "GG"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox TXTCFC 
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
         TabIndex        =   22
         Top             =   840
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox TXTEMAILC 
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
         TabIndex        =   17
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox TXTPREFISSOCELC 
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
         MaxLength       =   3
         TabIndex        =   16
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox TXTNAZIONALITAC 
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
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TXTPREFISSOTC 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox TXTCAPC 
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
         TabIndex        =   13
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TXTCITTAC 
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
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TXTPROVINCIAC 
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
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TXTVIAC 
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
         TabIndex        =   10
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TXTNCIVICOC 
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
         TabIndex        =   9
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox TXTCOGNOMEC 
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
         Left            =   3600
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TXTNOMEC 
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
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Frame FMETITOLO 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   4
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
            TabIndex        =   5
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
            Caption         =   "CANTANTE"
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
            TabIndex        =   6
            Top             =   120
            Width           =   1380
         End
      End
      Begin VB.Frame FMESESSO 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   1560
         Width           =   1215
         Begin VB.OptionButton OPTF 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "F"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   720
            TabIndex        =   3
            Top             =   120
            Width           =   375
         End
         Begin VB.OptionButton OPTM 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "M"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   43
         Top             =   2880
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
         NumItems        =   16
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CF"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NOME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "COGNOME"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DATA N."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SESSO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "VIA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "N"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PR"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "CITTA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "CAP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "NAZIONALITA"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "PREFT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "SUFT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "PREFC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "SUFC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "EMAIL"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label LBLSESSO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SESSO"
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
         Left            =   6120
         TabIndex        =   42
         Top             =   1320
         Width           =   855
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
         TabIndex        =   41
         Top             =   840
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
         Left            =   4560
         TabIndex        =   40
         Top             =   2040
         Width           =   1455
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
         Left            =   3000
         TabIndex        =   39
         Top             =   2040
         Width           =   1935
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
         Left            =   1680
         TabIndex        =   38
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label LBLNAZIONALITAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NAZIONALITA"
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
         TabIndex        =   37
         Top             =   2040
         Width           =   1335
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
         Left            =   4800
         TabIndex        =   36
         Top             =   1320
         Width           =   735
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
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LBLCITTAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CITTA'"
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
         TabIndex        =   34
         Top             =   1320
         Width           =   615
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
         TabIndex        =   33
         Top             =   1320
         Width           =   495
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
         TabIndex        =   32
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LBLDATANASCITAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DATA NASCITA"
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
         Left            =   5040
         TabIndex        =   31
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label LBLCOGNOMEC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "COGNOME"
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
         Left            =   3600
         TabIndex        =   30
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LBLNOMEC 
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
         Left            =   2400
         TabIndex        =   29
         Top             =   600
         Width           =   735
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
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   1320
   End
End
Attribute VB_Name = "FRMCANTANTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DB As Database
Dim RS As Recordset
Dim A As Integer

Public Function CONTROLLACODICE(CODCOM As String) As Boolean
CERCACODICE = "CF_C='" + CODCOM + "'"
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
Set LIST = ListView1.ListItems.Add(, , RS!CF_C)
LIST.SubItems(1) = RS!NOME_C
LIST.SubItems(2) = RS!COGNOME_C
LIST.SubItems(3) = RS!DATA_NASCITA_C
LIST.SubItems(4) = RS!SESSO_C
LIST.SubItems(5) = RS!VIA_C
LIST.SubItems(6) = RS!N_CIVICO_C
LIST.SubItems(7) = RS!CITTA_C
LIST.SubItems(8) = RS!PROVINCIA_C
LIST.SubItems(9) = RS!CAP_C
LIST.SubItems(10) = RS!NAZIONALITA_C
LIST.SubItems(11) = RS!PREFISSOT_C
LIST.SubItems(12) = RS!SUFFISSOT_C

LIST.SubItems(13) = RS!PREFISSOC_C

LIST.SubItems(14) = RS!SUFFISSOC_C

LIST.SubItems(15) = RS!E_MAIL_C
RS.MoveNext
Loop
End Sub

Sub PULISCI()
TXTCFC.Text = ""
TXTNOMEC.Text = ""
TXTCOGNOMEC.Text = ""
CBOGGC.Text = "GG"
CBOMMC.Text = "MM"
CBOYYC.Text = "YY"
TXTVIAC.Text = ""
TXTNCIVICOC.Text = ""
TXTCITTAC.Text = ""
TXTPROVINCIAC.Text = ""
TXTCAPC.Text = ""
TXTNAZIONALITAC.Text = ""
TXTPREFISSOTC.Text = ""
TXTSUFFISSOTC.Text = ""
TXTPREFISSOCELC.Text = ""
TXTSUFFISSOCELC.Text = ""
TXTEMAILC.Text = ""
OPTM.Value = False
OPTF.Value = False
End Sub
Sub TRASFERISCI()
RS("CF_C") = TXTCFC.Text
RS("NOME_C") = TXTNOMEC.Text
RS("COGNOME_C") = TXTCOGNOMEC.Text
RS("DATA_NASCITA_C") = CBOGGC.Text + "/" + CBOMMC.Text + "/" + CBOYYC.Text
If OPTM.Value = True Then
    RS("SESSO_C") = "MASCHIO"
ElseIf OPTF.Value = True Then
    RS("SESSO_C") = "FEMMINA"
End If
RS("VIA_C") = TXTVIAC.Text
RS("N_CIVICO_C") = TXTNCIVICOC.Text
RS("CITTA_C") = TXTCITTAC.Text
RS("PROVINCIA_C") = TXTPROVINCIAC.Text
RS("CAP_C") = Val(TXTCAPC.Text)
RS("NAZIONALITA_C") = TXTNAZIONALITAC.Text
If TXTPREFISSOTC.Text = "" Then
RS("PREFISSOT_C") = 0
Else
RS("PREFISSOT_C") = TXTPREFISSOTC.Text
End If
RS("SUFFISSOT_C") = TXTSUFFISSOTC.Text
RS("PREFISSOC_C") = TXTPREFISSOCELC.Text
RS("SUFFISSOC_C") = TXTSUFFISSOCELC.Text
RS("E_MAIL_C") = TXTEMAILC.Text
End Sub



Private Sub TXTCFC_LostFocus()
If CONTROLLACODICE(TXTCFC) = True Then
MsgBox ("CODISE FISCALE CLIENTE ESISTENTE")
TXTCFC.Text = ""
TXTCFC.SetFocus
End If
End Sub



Private Sub CMDBACK_Click()
A = 1
FRMHOME.Show

End Sub

Private Sub Form_ACTIVATE()
Timer1.Enabled = True
FRMCANTANTE.Top = 8000
FRMCANTANTE.Height = 0
A = 0
CARICADATA
Set DB = Workspaces(0).OpenDatabase(App.Path + "\DATABASE")
Set RS = DB.OpenRecordset("T_CANTANTI", dbOpenDynaset)
If RS.RecordCount <> 0 Then
CARICARECORD
End If
End Sub

Private Sub CMDSALVA_Click()
If Len(TXTCFC.Text) = 16 And TXTNOMEC.Text <> "" And TXTCOGNOMEC.Text <> "" And TXTVIAC.Text <> "" And TXTCAPC.Text <> "" And TXTCITTAC.Text <> "" And TXTPROVINCIAC.Text <> "" And TXTNAZIONALITAC.Text <> "" And TXTPREFISSOTC.Text <> "" And TXTSUFFISSOTC.Text <> "" And TXTSUFFISSOTC.Text <> "" And TXTSUFFISSOTC.Text <> "" Then
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
    ricercamodifica = "CF_C='" & TXTCFC.Text & "'"
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
    ricercacfc = "cf_c='" & TXTCFC.Text & "'"
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
    If FRMCANTANTE.Height > 4400 Then
        FRMCANTANTE.Height = 4890
    Else
        FRMCANTANTE.Height = FRMCANTANTE.Height + 200
        FRMCANTANTE.Top = FRMCANTANTE.Top - 200

    End If
ElseIf A = 1 Then
    If FRMCANTANTE.Height < 100 Then
        Unload Me
    Else
        FRMCANTANTE.Height = FRMCANTANTE.Height - 200
        FRMCANTANTE.Top = FRMCANTANTE.Top + 200
    End If

End If
End Sub


