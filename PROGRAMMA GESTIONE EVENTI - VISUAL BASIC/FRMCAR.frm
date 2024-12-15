VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMCAR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "CARS"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7065
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5459.384
   ScaleMode       =   0  'User
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FMESUPPORTO 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4905
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
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
         ItemData        =   "FRMCAR.frx":0000
         Left            =   3000
         List            =   "FRMCAR.frx":0002
         TabIndex        =   34
         Text            =   "GG"
         Top             =   2280
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
         Left            =   3600
         TabIndex        =   33
         Text            =   "MM"
         Top             =   2280
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
         Left            =   4200
         TabIndex        =   32
         Text            =   "YY"
         Top             =   2280
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   0
         Top             =   4920
      End
      Begin VB.TextBox TXTMARCA 
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
         TabIndex        =   19
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TXTTARGA 
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
         MaxLength       =   7
         TabIndex        =   18
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox TXTTIPO 
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
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox TXTTELAIO 
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox TXTPOSTI 
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
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TXTCOLORE 
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
         TabIndex        =   10
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TXTCAVALLI 
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
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TXTMODELLO 
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
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TXTKW 
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
         TabIndex        =   7
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox TXTMATRICOLA 
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
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.Frame FMETITOLO 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   120
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
            TabIndex        =   3
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
            Caption         =   "CAR"
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
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   480
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
      End
      Begin VB.TextBox TXTCILINDRATA 
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
         TabIndex        =   1
         Top             =   1560
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   36
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
      Begin VB.Label LBLDATAM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DATA MATRICOLA"
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
         TabIndex        =   35
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label LBLDEN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MARCA"
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
         TabIndex        =   31
         Top             =   600
         Width           =   1695
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
         TabIndex        =   30
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label LBLCELLULARE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TARGA"
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
         TabIndex        =   29
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label LBLTELEFONOC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO"
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
      Begin VB.Label LBLCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "POSTI"
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
         TabIndex        =   27
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label LBLPROVINCIAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CAVALLI"
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
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label LBLCITTAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "COLORE"
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
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label LBLNCIVICOC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "KW"
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
         TabIndex        =   24
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label LBLVIAC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MODELLO"
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
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LBLNOMEC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MATRICOLA"
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
         TabIndex        =   22
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label LBLCFC 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TELAIO"
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
         TabIndex        =   21
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label LBLSTATO 
         BackStyle       =   0  'Transparent
         Caption         =   "CILINDRATA"
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
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FRMCAR"
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

Private Sub CMDELIMINA_Click()
RISPOSTA = MsgBox("sei sicuro di voler eliminare il record", vbYesNo)
If RISPOSTA = 6 Then
    ricercacfc = "MATRICOLA='" & TXTMATRICOLA.Text & "'"
    RS.FindFirst (ricercacfc)
    If RS.NoMatch = False Then
    RS.Delete
    PULISCI
    CARICARECORD
    End If
End If
End Sub

Private Sub CMDMODIFICA_Click()
C = MsgBox("SEI SICURO DI VOLER MODIFICARE IL RECORD", vbYesNo)
If C = 6 Then
    ricercamodifica = "MATRICOLA='" & TXTMATRICOLA.Text & "'"
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

Private Sub CMDNUOVO_Click()
PULISCI
ListView1.Refresh
CMDSALVA.Enabled = True
CMDNUOVO.Enabled = False
CMDMODIFICA.Enabled = False
End Sub

Private Sub CMDSALVA_Click()
If TXTMATRICOLA.Text <> "" And TXTTELAIO.Text <> "" And TXTMARCA.Text <> "" And TXTMARCA.Text <> "" And TXTMODELLO.Text <> "" And TXTKW.Text <> "" And TXTCAVALLI.Text <> "" And TXTCILINDRATA.Text <> "" And TXTCOLORE.Text <> "" And TXTPOSTI.Text <> "" And TXTTIPO.Text <> "" And TXTTARGA.Text <> "" Then
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
End Sub
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
    If FRMCAR.Height > 4400 Then
        FRMCAR.Height = 4890
    Else
        FRMCAR.Height = FRMCAR.Height + 200
        FRMCAR.Top = FRMCAR.Top - 200

    End If
ElseIf A = 1 Then
    If FRMCAR.Height < 100 Then
        Timer1.Enabled = False
        Unload Me
    Else
        FRMCAR.Height = FRMCAR.Height - 200
        FRMCAR.Top = FRMCAR.Top + 200
    End If

End If
End Sub
Private Sub Form_activate()
Timer1.Enabled = True
FRMCAR.Top = 8000
FRMCAR.Height = 0
A = 0
Set DB = Workspaces(0).OpenDatabase(App.Path + "\DATABASE")
Set RS = DB.OpenRecordset("T_CAR", dbOpenDynaset)
If RS.RecordCount <> 0 Then
CARICARECORD
End If
End Sub
Private Sub TXTMATRICOLA_LostFocus()
If CONTROLLACODICE(TXTMATRICOLA) = True Then
MsgBox ("MATRICOLA ESISTENTE")
TXTMATRICOLA.Text = ""
TXTMATRICOLA.SetFocus
End If
End Sub
Sub TRASFERISCI()
RS("MATRICOLA") = TXTMATRICOLA.Text
RS("TELAIO") = TXTTELAIO.Text
RS("MARCA") = TXTCOGNOMEC.Text
RS("MODELLO") = TXTMODELLO.Text
RS("KW") = TXTKW.Text
RS("CAVALLI") = TXTCAVALLI.Text
RS("CILINDRATA") = TXTCILINDRATA.Text
RS("COLORE") = Val(TXTCOLORE.Text)
RS("POSTI") = TXTPOSTI.Text
RS("TARGA") = TXTTARGA.Text
RS("DATA_NASCITA_C") = CBOGGC.Text + "/" + CBOMMC.Text + "/" + CBOYYC.Text
End Sub
Sub PULISCI()
TXTMATRICOLA.Text = ""
TXTTELAIO.Text = ""
TXTMARCA.Text = ""
TXTMODELLO.Text = ""
TXTKW.Text = ""
TXTCAVALLI.Text = ""
TXTCILINDRATA.Text = ""
TXTCOLORE.Text = ""
TXTPOSTI.Text = ""
TXTTIPO.Text = ""
TXTTARGA.Text = ""
CBOGGC.Text = "GG"
CBOMMC.Text = "MM"
CBOYYC.Text = "YY"
End Sub
Public Function CONTROLLAMATRICOLA(CODCOM As String) As Boolean
CERCAMATRICOLA = "MATRICOLA='" + CODCOM + "'"
RS.FindFirst (CERCAMATRICOLA)
If RS.NoMatch = True Then
CONTROLLAMATRICOLA = False
Else
CONTROLLAMATRICOLA = True
Exit Function
End If
End Function
Sub CARICARECORD()
Dim LIST As ListItem
ListView1.ListItems.Clear
RS.MoveFirst
Do Until RS.EOF
Set LIST = ListView1.ListItems.Add(, , RS!MATRICOLA)
LIST.SubItems(1) = RS!TELAIO
LIST.SubItems(2) = RS!MARCA
LIST.SubItems(3) = RS!MODELLO
LIST.SubItems(4) = RS!KW
LIST.SubItems(5) = RS!CAVALLI
LIST.SubItems(6) = RS!CILINDRATA
LIST.SubItems(7) = RS!COLORE
LIST.SubItems(8) = RS!POSTI
LIST.SubItems(9) = RS!TIPO
LIST.SubItems(10) = RS!TARGA
LIST.SubItems(11) = RS!DATA_NASCITA
RS.MoveNext
Loop
End Sub
