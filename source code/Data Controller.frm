VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DataBase Controller - ( P.SelvaKumar )"
   ClientHeight    =   8220
   ClientLeft      =   3090
   ClientTop       =   1110
   ClientWidth     =   10320
   Icon            =   "Data Controller.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10320
   Begin VB.CommandButton Command6 
      Caption         =   "Save As &Text File"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   43
      ToolTipText     =   "Save the Current Record as Text File"
      Top             =   7440
      Width           =   1695
   End
   Begin ComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   42
      Top             =   7920
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13018
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "4/26/2007"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "4:32 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   12
      ToolTipText     =   "Close Application"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "A&bout"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      ToolTipText     =   "About the Creator"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete Current Record"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3600
      TabIndex        =   10
      ToolTipText     =   "Delete the Current Record"
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Edit Current Record"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      TabIndex        =   9
      ToolTipText     =   "Edit the Current Record"
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Add New Record"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9000
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6960
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4210
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   6960
      Width           =   4800
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   6960
      Width           =   4215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5460
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   10335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select a Database to Search"
      Top             =   120
      Width           =   2535
   End
   Begin VB.ComboBox search 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Enter a Software name to Search"
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label Label31 
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   41
      Top             =   6585
      Width           =   735
   End
   Begin VB.Label Label30 
      Caption         =   "Serial No / Crack Path / Software In"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   40
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label Label29 
      Caption         =   "Software Name"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   39
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   26
      Left            =   9980
      MouseIcon       =   "Data Controller.frx":27C92
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   25
      Left            =   9650
      MouseIcon       =   "Data Controller.frx":27F9C
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   24
      Left            =   9300
      MouseIcon       =   "Data Controller.frx":282A6
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   23
      Left            =   8900
      MouseIcon       =   "Data Controller.frx":285B0
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Top             =   660
      Width           =   240
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   22
      Left            =   8550
      MouseIcon       =   "Data Controller.frx":288BA
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   21
      Left            =   8200
      MouseIcon       =   "Data Controller.frx":28BC4
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   20
      Left            =   7880
      MouseIcon       =   "Data Controller.frx":28ECE
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   19
      Left            =   7550
      MouseIcon       =   "Data Controller.frx":291D8
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   18
      Left            =   7160
      MouseIcon       =   "Data Controller.frx":294E2
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   17
      Left            =   6800
      MouseIcon       =   "Data Controller.frx":297EC
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   16
      Left            =   6480
      MouseIcon       =   "Data Controller.frx":29AF6
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   15
      Left            =   6110
      MouseIcon       =   "Data Controller.frx":29E00
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   14
      Left            =   5710
      MouseIcon       =   "Data Controller.frx":2A10A
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   13
      Left            =   5350
      MouseIcon       =   "Data Controller.frx":2A414
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   660
      Width           =   210
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   12
      Left            =   5040
      MouseIcon       =   "Data Controller.frx":2A71E
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   660
      Width           =   240
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   11
      Left            =   4650
      MouseIcon       =   "Data Controller.frx":2AA28
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   4350
      MouseIcon       =   "Data Controller.frx":2AD32
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   4000
      MouseIcon       =   "Data Controller.frx":2B03C
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   8
      Left            =   3620
      MouseIcon       =   "Data Controller.frx":2B346
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   660
      Width           =   165
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   3250
      MouseIcon       =   "Data Controller.frx":2B650
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   2930
      MouseIcon       =   "Data Controller.frx":2B95A
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   2550
      MouseIcon       =   "Data Controller.frx":2BC64
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   2200
      MouseIcon       =   "Data Controller.frx":2BF6E
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   1840
      MouseIcon       =   "Data Controller.frx":2C278
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   1500
      MouseIcon       =   "Data Controller.frx":2C582
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   660
      Width           =   135
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   195
      Index           =   1
      Left            =   1150
      MouseIcon       =   "Data Controller.frx":2C88C
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   660
      Width           =   120
   End
   Begin VB.Label lab 
      BackStyle       =   0  'Transparent
      Caption         =   "Show All"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   240
      MouseIcon       =   "Data Controller.frx":2CB96
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   675
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   75
      Picture         =   "Data Controller.frx":2CEA0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   10185
   End
   Begin VB.Label Label1 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As String
Dim sql As String
Dim tmp As Integer
Dim ins As String
Dim aa As Boolean
Public Sub tru()
    Dim i As Integer
        Text1.Locked = True
        Text2.MaxLength = 0
        Text2.Locked = True
        Combo2.Locked = True
        search.Enabled = True
        Combo1.Enabled = True
        List1.Enabled = True
        Command3.Enabled = True
        Command6.Enabled = True
        For i = 0 To 26
            lab(i).Enabled = True
        Next i
        Set rs = New ADODB.Recordset
        rs.Open sql, cn, adOpenKeyset
        Call dis
End Sub
Public Sub fal()
Dim i As Integer
    Text1.Locked = False
    Text2.Locked = False
    Text2.MaxLength = 120
    Combo2.Locked = False
    search.Enabled = False
    Combo1.Enabled = False
    List1.Enabled = False
    Command3.Enabled = False
    Command6.Enabled = False
    For i = 0 To 26
        lab(i).Enabled = False
    Next i
End Sub
Private Sub Combo1_Click()
If tmp <> 27 Then
    Call lab_Click(tmp)
Else
    Call search_KeyPress(13)
End If
End Sub
Private Sub Command1_Click()
If Command1.Caption = "&Add New" Then
    Command1.Caption = "&Save"
    Command2.Caption = "&Cancel Add New"
    Call fal
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
ElseIf Command2.Caption = "&Cancel Add New" Then
    If Len(Trim(Text1.Text)) <> 0 And Len(Trim(Text2.Text)) <> 0 Then
        ins = "insert into keys(name,key,type) values('" & Trim(Text1.Text) & "','" & Trim(Text2.Text) & "'," & Combo2.ListIndex & ")"
        cn.Execute ins
        Command1.Caption = "&Add New"
        Command2.Caption = "&Edit Current Record"
        Call tru
        MsgBox "Sucessfully Saved in the Database ! ", vbInformation
    Else
        MsgBox "Please Input Values ! ", vbInformation
    End If
Else
    If Len(Trim(Text1.Text)) <> 0 And Len(Trim(Text2.Text)) <> 0 Then
        ins = "update keys set name='" & Trim(Text1.Text) & "',key='" & Trim(Text2.Text) & "',type=" & Combo2.ListIndex & " where sl_no=" & rs("sl_no")
        cn.Execute ins
        Command1.Caption = "&Add New"
        Command2.Caption = "&Edit Current Record"
        Call tru
        MsgBox "Sucessfully Saved in the Database ! ", vbInformation
     Else
        MsgBox "Please Enter Values ! ", vbInformation
    End If
End If
End Sub
Private Sub Command2_Click()
If Command2.Caption = "&Edit Current Record" Then
    If rs.RecordCount <> 0 Then
        Text2.Text = rs("key")
        Command1.Caption = "&Save"
        Command2.Caption = "&Cancel Edit"
        Call fal
        Text1.SetFocus
    Else
        MsgBox "Invalid Record Selection, Nothing to Edit ! ", vbInformation
    End If
Else
    Command1.Caption = "&Add New"
    Command2.Caption = "&Edit Current Record"
    Call tru
    search.SetFocus
End If
End Sub
Private Sub Command3_Click()
Dim a As Integer
If rs.RecordCount <> 0 Then
a = MsgBox("Are you Sure want to Delete Current Record ? ", vbInformation + vbYesNo)
If a = 6 Then
    ins = "delete from keys where sl_no=" & rs("sl_no")
    cn.Execute ins
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenKeyset
    Call dis
    MsgBox "Sucessfully Deleted  ", vbInformation
End If
Else
MsgBox "Invalid Record Selection, Nothing to Delete ! ", vbInformation
End If
End Sub
Private Sub Command4_Click()
MsgBox "This Application has been Designed and Programmed by : P.SelvaKumar   ", vbInformation, Form1.Caption
End Sub
Private Sub Command5_Click()
Unload Me
End Sub
Private Sub Command6_Click()
If rs.RecordCount <> 0 Then
Dim i As Integer
List1.Enabled = False
filenum = FreeFile
rs.MoveFirst
    Open "C:\Documents and Settings\All Users\Desktop\PSSK.txt" For Output As filenum
    For i = 1 To rs.RecordCount
        Select Case rs("type")
            Case 0:
                Print #filenum, i & ". " & rs("name") & ", Type : Software"
            Case 1:
                Print #filenum, i & ". " & rs("name") & ", Type : Game"
            Case 2:
                Print #filenum, i & ". " & rs("name") & ", S/n : " & rs("key")
            Case 3:
                Print #filenum, i & ". " & rs("name")
        End Select
        rs.MoveNext
    Next i
    Close filenum
    List1.Enabled = True
    MsgBox "Text File (PSSK.txt) Sucessfully Created in Desktop ! ", vbInformation
Else
    MsgBox "Are You Kiting Me ? , Nothing to Write ! ", vbInformation
End If
End Sub
Private Sub Form_Load()
On Error GoTo myerr
Combo2.AddItem "Software"
Combo2.AddItem "Game"
Combo2.AddItem "Serial No"
Combo2.AddItem "Crack"
Combo1.AddItem "Software Database"
Combo1.AddItem "Game Database"
Combo1.AddItem "Serial Database"
Combo1.AddItem "Crack Database"
Combo1.AddItem "Both Serial & Crack Database"
cmd = "Provider=microsoft.jet.oledb.3.51;data source=" + App.Path + "\Support\PSSK.mdb"
Set cn = New ADODB.Connection
With cn
    .ConnectionString = cmd
    .Open
End With
myerr:
    If Err.Number = -2147467259 Then
        FileCopy App.Path + "\Support\PSSK.mdb", "C:\PSSK.mdb"
        cmd = "Provider=microsoft.jet.oledb.3.51;data source=C:\PSSK.mdb"
        Set cn = New ADODB.Connection
        With cn
            .ConnectionString = cmd
            .Open
        End With
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
    End If
    tmp = 1
Combo1.ListIndex = 4
End Sub
Private Sub Form_Unload(Cancel As Integer)
rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing
If cmd = "Provider=microsoft.jet.oledb.3.51;data source=C:\PSSK.mdb" Then
    Kill "C:\PSSK.mdb"
End If
End Sub
Private Sub lab_Click(Index As Integer)
On Error Resume Next
    sql = "select * from keys"
    If Index <> 0 Then
    sql = sql & " where name like '" & lab(Index).Caption & "%'"
    End If
    lab(Index).ForeColor = &H40C0&
    If tmp <> 27 And tmp <> Index Then
        lab(tmp).ForeColor = vbBlue
    End If
    tmp = Index
    If Index = 0 And Combo1.ListIndex = 4 Then
        sql = sql & " where type=2 or type=3 order by name"
    ElseIf Index = 0 Then
        sql = sql & " where type=" & Combo1.ListIndex & " order by name"
    ElseIf Combo1.ListIndex = 4 Then
        sql = sql & " and (type=2 or type=3) order by name"
    Else
        sql = sql & " and type=" & Combo1.ListIndex & " order by name"
    End If
Set rs = New ADODB.Recordset
rs.Open sql, cn, adOpenKeyset
Call dis
End Sub
Public Sub dis()
status.Panels(1).Text = "Searching Please Wait ...."
List1.Clear
If rs.RecordCount <> 0 Then
Dim i As Integer
rs.MoveFirst
For i = 1 To rs.RecordCount
    List1.AddItem rs("name")
    rs.MoveNext
Next i
Else
List1.AddItem "# NO MATCH FOUND IN THE DATABASE"
Text1.Text = ""
Text2.Text = ""
End If
status.Panels(1).Text = rs.RecordCount & "  Records Found"
List1.Selected(0) = True
End Sub
Private Sub List1_Click()
If rs.RecordCount <> 0 Then
rs.Move List1.ListIndex, 1
Text1.Text = rs("name")
If Val(rs("type")) = 3 Then
    Text2.Text = App.Path & "\Crack Collection\" & rs("key")
Else
    Text2.Text = rs("key")
End If
Combo2.ListIndex = Val(rs("type"))
End If
End Sub
Private Sub search_GotFocus()
SendKeys "{Home}+{End}"
aa = True
End Sub
Private Sub search_KeyPress(KeyAscii As Integer)
On Error GoTo sqlerr
If KeyAscii = 13 And Len(Trim(search.Text)) = 0 Then
    search.Text = Trim(search.Text)
    MsgBox "Please Enter a Software Name to Search ! ", vbInformation
End If
If KeyAscii = 13 And Len(Trim(search.Text)) <> 0 Then
Dim i As Integer
sql = ""
For i = 1 To Len(Trim(search.Text))
    If Mid(Trim(search.Text), i, 1) = "%" Then
            sql = Trim(search.Text)
            Exit For
    Else
            sql = sql + "%" + Mid(Trim(search.Text), i, 1)
    End If
Next i
sql = sql + "%"
If Combo1.ListIndex = 4 Then
sql = "select * from keys where name like '" + sql + "' and (type=2 or type=3) order by name"
Else
sql = "select * from keys where name like '" + sql + "' and type=" & Combo1.ListIndex & " order by name"
End If
Set rs = New ADODB.Recordset
rs.Open sql, cn, adOpenKeyset
search.AddItem search.Text
If Val(tmp) <> 27 Then
lab(tmp).ForeColor = vbBlue
tmp = 27
End If
Call dis
If aa = True Then
SendKeys "{Home}+{End}"
End If
End If
If KeyAscii = 10 Then
    If Len(search.Text) = 0 Then
        MsgBox "Enter a SQL Query to Run ! ", vbInformation
    ElseIf Mid(LCase(Trim(search.Text)), 1, 6) = "select" Then
        Set rs = New ADODB.Recordset
        rs.Open Trim(search.Text), cn, adOpenKeyset
        Call dis
    Else
        cn.Execute (Trim(search.Text))
        Call Combo1_Click
    End If
    If aa = True Then
    SendKeys "{Home}+{End}"
    End If
End If
sqlerr:
    If Err.Number <> 0 Then
        MsgBox "Invalid SQL Query ! ", vbInformation
        If KeyAscii = 13 Then
                Call lab_Click(1)
        Else
                Call Combo1_Click
        End If
    End If
End Sub
Private Sub search_LostFocus()
aa = False
End Sub
Private Sub Text1_GotFocus()
SendKeys "{Home}+{End}"
End Sub
Private Sub Text2_GotFocus()
SendKeys "{Home}+{End}"
End Sub
