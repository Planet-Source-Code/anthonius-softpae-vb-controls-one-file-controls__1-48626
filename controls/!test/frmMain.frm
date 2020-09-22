VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Test - Flat Control"
   ClientHeight    =   3390
   ClientLeft      =   3810
   ClientTop       =   1785
   ClientWidth     =   6465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6465
   Begin Project1.List List1 
      Height          =   2310
      Left            =   3030
      TabIndex        =   0
      Top             =   825
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   4075
   End
   Begin Project1.Tree Tree1 
      Height          =   2160
      Left            =   180
      TabIndex        =   3
      Top             =   825
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   3810
      Border          =   1
      Indent          =   16
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3015
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   405
      Width           =   3270
   End
   Begin Project1.Flater Flater1 
      Left            =   3480
      Top             =   -45
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox txtInPicture 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   420
      Width           =   2745
   End
   Begin Project1.Flater Flater2 
      Left            =   3015
      Top             =   -60
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin Project1.ButtonXP ButtonXP2 
      Height          =   330
      Left            =   705
      TabIndex        =   4
      Top             =   30
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":1272
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP1 
      Height          =   330
      Left            =   180
      TabIndex        =   5
      Top             =   30
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":13CC
      Style           =   2
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP3 
      Height          =   330
      Left            =   1050
      TabIndex        =   6
      Top             =   30
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":1526
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.softpae.com"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4740
      TabIndex        =   7
      Top             =   75
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFlat1 As FlatCtl
Dim cFlat2 As FlatCtl

Private Sub ButtonXP1_Click()
    MsgBox "Button Clicked !"
End Sub

Private Sub ButtonXP1_MenuClick()
    MsgBox "MenuButton Clicked !"
End Sub

Private Sub Form_Load()

    Set cFlat1 = New FlatCtl
    Set cFlat2 = New FlatCtl
    
    ' You can use control of class :))
    'cFlat1.Attach txtInPicture
    'cFlat2.Attach Combo1
    
    Flater1.Attach txtInPicture
    Flater2.Attach Combo1
    
    Tree1.hImageList = Tree1.LoadList(App.Path & "\resource.bmp", &H0, False, 752, 16)
    
    Call Tree1.AddItem("", 0, "root", "Zoznam", 0)
    Call Tree1.AddItem("root", tvwFirst, "today", "Private", 1)
    Call Tree1.AddItem("today", tvwLast, "notes", "Poznámky", 2)
    Call Tree1.AddItem("today", tvwLast, "calendar", "Kalendár", 3)
    Call Tree1.AddItem("today", tvwLast, "contacts", "Adresár", 4)
    Call Tree1.AddItem("today", tvwLast, "favorites", "Ob¾úbené", 5)
    
    List1.ListImage = "bmp_01"
    For k = 1 To Screen.FontCount
        List1.AddItem Screen.Fonts(k - 1)
    Next
    
    Me.Show
    
End Sub
