VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#17.2#0"; "Codejock.SkinFramework.v17.2.0.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.CommandBars.v17.2.0.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "Codejock.DockingPane.v17.2.0.ocx"
Begin VB.Form frmppal 
   Caption         =   "Ariconta"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11580
   FillStyle       =   0  'Solid
   Icon            =   "frmPpalN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8880
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1856A
            Key             =   "New"
            Object.Tag             =   "100"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":185C8
            Key             =   "Open"
            Object.Tag             =   "101"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18626
            Key             =   "Save"
            Object.Tag             =   "103"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18684
            Key             =   "Print"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":186E2
            Key             =   "Cut"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18740
            Key             =   "Copy"
            Object.Tag             =   "106"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1879E
            Key             =   "Paste"
            Object.Tag             =   "107"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":187FC
            Key             =   "Bold"
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1885A
            Key             =   "Italic"
            Object.Tag             =   "121"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":188B8
            Key             =   "Underline"
            Object.Tag             =   "122"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18916
            Key             =   "Align Left"
            Object.Tag             =   "123"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18974
            Key             =   "Center"
            Object.Tag             =   "124"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":189D2
            Key             =   "Align Right"
            Object.Tag             =   "125"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18A30
            Key             =   "About"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18A8E
            Key             =   ""
            Object.Tag             =   "166"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18AEC
            Key             =   ""
            Object.Tag             =   "168"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18B4A
            Key             =   ""
            Object.Tag             =   "165"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListPPal48 
      Left            =   5280
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   1440
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   1680
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   2040
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   2400
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageListPpal16 
      Left            =   360
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgListComun 
      Left            =   1920
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32 
      Left            =   2400
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":18BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1F40A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":25C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":2C4CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":32D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":39592
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":3FDF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras 
      Left            =   2880
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":46656
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":4CEB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":5371A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":59F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":607DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":67040
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":6D8A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":74104
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras_BN 
      Left            =   2760
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483626
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":74B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7B378
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":81BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":8843C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":8EC9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":95500
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":9BD62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":A25C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":A8E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":AF688
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2880
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B009A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B68FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B90AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListDocumentos 
      Left            =   2400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BF910
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":C0B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":C3344
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":C547E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":C5798
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":C8B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":CA79C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":CB579
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":CC4E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":CD460
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32_BN 
      Left            =   3120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":CE3FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":D4C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":DB4C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":E1D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":E8585
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":EEDE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":F5649
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListviews 
      Left            =   2880
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":FBEAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":10270D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":104EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":10AAE1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcoForms 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":111343
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":111D55
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":111DF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   1200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5640
      Top             =   1080
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   4800
      Top             =   1920
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":112802
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   3840
      Top             =   600
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   4320
      Top             =   1320
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager ImageManagerGalleryStyles 
      Left            =   3360
      Top             =   120
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":11281C
   End
End
Attribute VB_Name = "frmppal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Dim ContextEvent As CalendarEvent

'Public DisableDragging_ForRecurrenceEvents As Boolean
'Public DisableInPlaceCreateEvents_ForSaSu As Boolean
'
'Public EnableScrollV_DayView As Boolean
'Public EnableScrollH_DayView As Boolean
'
'Public EnableScrollV_WeekView As Boolean
'
'Public EnableScrollV_MonthView As Boolean
'Public ToolTips_Mode As Long
'
'Dim mailIconArray(0 To 9) As Long
'Dim toolbarIconArray(0 To 51) As Long
'
'
'
'
'
'
'Dim WithEvents GalleryQSItems As CommandBarGalleryItems

Dim MRUShortcutBarWidth


Const IMAGEBASE = 10000
Const MinimizedShortcutBarWidth = 32 + 8

Dim WithEvents statusBar  As XtremeCommandBars.statusBar
Attribute statusBar.VB_VarHelpID = -1
Dim FontSizes(4) As Integer
Dim RibbonSeHaCreado As Boolean
Dim Pane As Pane
Dim Cad As String

'Variables comunes para todos los procedimientos de carga menus en el ribbon
'Codejock
Dim TabNuevo As RibbonTab
Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup
Dim GroupManageCalendars As RibbonGroup, GroupShare As RibbonGroup, GroupFind As RibbonGroup
Dim idTabPpal As Integer

Dim Control As CommandBarControl
Dim ControlNew_NewItems As CommandBarPopup
Dim Rn2 As ADODB.Recordset
Dim Habilitado As Boolean


Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
End Function

Sub LoadResources(DllName As String, IniFileName As String)
Dim elpath As String
    
      elpath = App.Path & "\Styles\"
    CommandBarsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ShortcutBarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    SuiteControlsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    CalendarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ReportControlGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    DockingPaneGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
End Sub

Public Sub CheckButton(nButton As Integer)
    CommandBars.Actions(ID_OPTIONS_STYLEBLUE2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLESILVER2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLEBLACK2010).Checked = False
    
    CommandBars.Actions(nButton).Checked = True
End Sub

Sub OnThemeChanged(id As Integer)
Dim N_Skin As Integer
    CheckButton id
    
    Dim FlatStyle As Boolean
    FlatStyle = id >= ID_OPTIONS_STYLESCENIC7 And id <= ID_OPTIONS_STYLEBLACK2010
        
        
    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
   
    
    CommandBars.EnableOffice2007Frame False

    Select Case CommandBars.VisualTheme
        Case xtpThemeResource, xtpThemeRibbon
            CommandBars.AllowFrameTransparency False 'True
            CommandBars.EnableOffice2007Frame True
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
        Case Else
            CommandBars.AllowFrameTransparency True
            CommandBars.EnableOffice2007Frame False
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
    End Select
    
    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = CommandBars.ToolTipContext
    ToolTipContext.Style = xtpToolTipResource
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    
    statusBar.ToolTipContext.Style = ToolTipContext.Style
    frmShortBar.wndShortcutBar.ToolTipContext.Style = ToolTipContext.Style
    
       
    'CreateBackstage
    'SetBackstageTheme
    
    'CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\styles\Ariconta.ico", 16, 16
            
    'Set Captions VisualTheme
    On Error Resume Next
    Dim CtrlCaption As ShortcutCaption
    Dim Form As Form, Ctrl As Object
            
    For Each Form In Forms
        For Each Ctrl In Form.Controls
                    
            Set CtrlCaption = Ctrl
            If Not CtrlCaption Is Nothing Then
                CtrlCaption.VisualTheme = frmShortBar.wndShortcutBar.VisualTheme
            End If
                    
        Next
    Next
       
    DockingPaneManager.PaintManager.SplitterSize = 5
    DockingPaneManager.PaintManager.SplitterColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
    
    DockingPaneManager.PaintManager.ShowCaption = False
    DockingPaneManager.RedrawPanes
        
    frmShortBar.SetColor
    frmInbox.SetColor id
 

    frmPaneCalendar.SetFlatStyle FlatStyle
    frmPaneContacts.SetFlatStyle FlatStyle
    'frmPaneInformacion.SetFlatStyle FlatStyle
    'frmPaneAcercaDe.SetFlatStyle FlatStyle
    
    
    
    
    
    
    LoadIcons
    N_Skin = id - 2895
    EstablecerSkin N_Skin
    
    'Updatear SKIN usuario
    If CStr(N_Skin) <> vUsu.Skin Then
        vUsu.Skin = N_Skin
        vUsu.ActualizarSkin
    End If
    
End Sub

Public Sub SetBackstageTheme()
Dim I As Integer
    Dim nTheme As XtremeCommandBars.XTPBackstageButtonControlAppearanceStyle
    nTheme = xtpAppearanceResource

    If Not (pageBackstageInfo Is Nothing) Then
        pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        pageBackstageInfo.btnCheckForIssues.Appearance = nTheme
        pageBackstageInfo.btnManageVersions.Appearance = nTheme
    End If
    
    If Not (pageBackstageHelp Is Nothing) Then
        For I = 0 To 4
            pageBackstageHelp.btnAcciones(I).Appearance = nTheme
        Next
        
    End If
    
    If Not (pageBackstageSend Is Nothing) Then
        'pageBackstageSend.btnTab(0).Appearance = nTheme
        'pageBackstageSend.btnTab(1).Appearance = nTheme
        'pageBackstageSend.btnTab(2).Appearance = nTheme
        'pageBackstageSend.btnTab(3).Appearance = nTheme
    End If

End Sub

Private Sub CreateStatusBar()
   
    If RibbonSeHaCreado Then
        'StatusBar.Pane(0).Value = vEmpresa.nomempre & "    " & vUsu.Login
        statusBar.Pane(0).Text = "Nº " & vEmpresa.codempre
        statusBar.Pane(1).Text = vEmpresa.nomempre
    
    Else
    
     Dim Pane As StatusBarPane
     Set statusBar = Nothing
     
     Set statusBar = CommandBars.statusBar
     statusBar.Visible = True
     
     
     Set Pane = statusBar.AddPane(ID_INDICATOR_PAGENUMBER)
     Pane.Text = "Nº " & vEmpresa.codempre
     Pane.Caption = "&C"
     Pane.Value = vEmpresa.nomempre & "    " & vUsu.Login
     Pane.Button = True
     Pane.SetPadding 8, 0, 8, 0
     
     Set Pane = statusBar.AddPane(ID_INDICATOR_WORDCOUNT)
     Pane.Text = vEmpresa.nomempre
     Pane.Caption = "&Word Count"
     Pane.Value = "1"
     Pane.Button = True
     Pane.SetPadding 8, 0, 8, 0
     
     
     Set Pane = statusBar.AddPane(0)
     Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
     Pane.BeginGroup = True
             
    '
     statusBar.RibbonDividerIndex = 3
     statusBar.EnableCustomization True
     
     CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowNever
     CommandBars.Options.ShowKeyboardTips = True
     CommandBars.Options.ToolBarAccelTips = True
    End If
End Sub

Private Sub DockBarRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.RecalcLayout
    BarOnLeft.GetWindowRect Left, top, Right, Bottom
    
    CommandBars.DockToolBar BarToDock, Right, (Bottom + top) / 2, BarOnLeft.Position

End Sub

Public Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim AbiertoFormulario  As Boolean
    AbiertoFormulario = False
    
    
    
    Select Case Control.id
        Case XTPCommandBarsSpecialCommands.XTP_ID_RIBBONCONTROLTAB:
           ' Debug.Print "Selected Tab is Changed"
        
          
        Case XTP_ID_RIBBONCUSTOMIZE:
            CommandBars.ShowCustomizeDialog 3
            
        Case ID_APP_ABOUT:
          
           LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "AriCONTA-6.html?"
   
        
        Case ID_FILE_NEW:
            'frmEmail.Show 0, Me
        
        
        
        Case ID_Licencia_Usuario_Final_txt, ID_Licencia_Usuario_Final_web, ID_Ver_Version_operativa_web
            OpcionesMenuInformacion Control.id
        
        
        
        Case ID_VIEW_STATUSBAR:
            CommandBars.statusBar.Visible = Not CommandBars.statusBar.Visible
            CommandBars.RecalcLayout
            
        Case ID_RIBBON_EXPAND:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
            Dim newFontHeight As Integer
            newFontHeight = FontSizes(Control.id - ID_OPTIONS_FONT_SYSTEM)
            RibbonBar.FontHeight = newFontHeight
            
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
            CommandBars.PaintManager.AutoResizeIcons = Not CommandBars.PaintManager.AutoResizeIcons
            CommandBars.RecalcLayout
            RibbonBar.RedrawBar
            
        Case ID_OPTIONS_STYLEBLUE2010:
            LoadResources "Office2010.dll", "Office2010Blue.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.wndCalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLUE2010
            
            
            
       Case ID_OPTIONS_STYLESILVER2010:
            LoadResources "Office2010.dll", "Office2010Silver.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.wndCalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLESILVER2010
        
       Case ID_OPTIONS_STYLEBLACK2010:
            LoadResources "Office2010.dll", "Office2010Black.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.wndCalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLACK2010
        
        Case ID_APP_EXIT:
            Unload Me
        
    
            
        Case ID_GROUP_GOTO_TODAY:
            Select Case frmInbox.wndCalendarControl.ViewType
                Case xtpCalendarDayView:
                    frmInbox.wndCalendarControl.DayView.ShowDay DateTime.Now, True
            
                Case xtpCalendarWorkWeekView:
                    frmInbox.wndCalendarControl.DayView.SetSelection DateTime.Now, DateTime.Now, True
                    frmInbox.wndCalendarControl.RedrawControl
            
                Case xtpCalendarWeekView:
                    frmInbox.wndCalendarControl.WeekView.SetSelection DateTime.Now, DateTime.Now, True
            
                Case xtpCalendarMonthView:
                    frmInbox.wndCalendarControl.MonthView.SetSelection DateTime.Now, DateTime.Now, True
            End Select
            
        Case ID_GROUP_GOTO_NEXT7DAYS:
            Dim lastDate As Date
            lastDate = frmInbox.wndCalendarControl.DayView.Days(frmInbox.wndCalendarControl.DayView.DaysCount - 1).Date
            frmInbox.wndCalendarControl.ViewType = xtpCalendarDayView
            frmInbox.wndCalendarControl.DayView.ShowDays lastDate + 1, lastDate + 7
            
        Case ID_GROUP_ARRANGE_DAY:
            frmInbox.wndCalendarControl.ViewType = xtpCalendarDayView
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            frmInbox.wndCalendarControl.ViewType = xtpCalendarWorkWeekView
            
        Case ID_GROUP_ARRANGE_WEEK:
            frmInbox.wndCalendarControl.UseMultiColumnWeekMode = True
            frmInbox.wndCalendarControl.ViewType = xtpCalendarWeekView

        Case ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_MONTH_LOW, _
             ID_GROUP_ARRANGE_MONTH_MEDIUM, ID_GROUP_ARRANGE_MONTH_HIGH:
            frmInbox.wndCalendarControl.ViewType = xtpCalendarMonthView
            
        Case ID_CALENDAREVENT_OPEN:
            frmInbox.mnuOpenEvent
            
        Case ID_CALENDAREVENT_DELETE:
            frmInbox.mnuDeleteEvent
            
        Case ID_CALENDAREVENT_NEW, ID_GROUP_NEW_APPOINTMENT:
            'falta### frmEditEvent.AllDayOverride = False
            frmInbox.mnuNewEvent
            frmInbox.wndCalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_MEETING:
            'falta### frmEditEvent.AllDayOverride = False
            'falta### frmEditEvent.chkMeeting.Value = 1
            frmInbox.mnuNewEvent
            frmInbox.wndCalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_ALLDAY:
            'falta### frmEditEvent.AllDayOverride = True
            frmInbox.mnuNewEvent
            frmInbox.wndCalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_CALENDAREVENT_CHANGE_TIMEZONE:
            frmInbox.mnuChangeTimeZone
            
        Case ID_CALENDAREVENT_60:
            frmInbox.mnuTimeScale 60
            
        Case ID_CALENDAREVENT_30:
            frmInbox.mnuTimeScale 30
            
        Case ID_CALENDAREVENT_15:
            frmInbox.mnuTimeScale 15
            
        Case ID_CALENDAREVENT_10:
            frmInbox.mnuTimeScale 10
            
        Case ID_CALENDAREVENT_5:
            frmInbox.mnuTimeScale 5
            
            
            
     
        Case Else
            AbiertoFormulario = True
            AbrirFormularios Control.id
            
            
    End Select
    
    
    If AbiertoFormulario Then
        AbiertoFormulario = False
        'mOTIVO... no lo se
        'Pero si lo vamos cambiando funciona
        If Me.DockingPaneManager.Panes(1).Enabled = 3 Then
            Me.DockingPaneManager.Panes(1).Enabled = 3
            Me.DockingPaneManager.Panes(2).Enabled = 3

            frmPaneCalendar.DatePicker.Enabled = True
            
            DockingPaneManager.RedrawPanes
            
            
        Else
            Me.DockingPaneManager.Panes(1).Enabled = 3
            Me.DockingPaneManager.Panes(2).Enabled = 3
             
        End If
        DockingPaneManager.NormalizeSplitters

    End If
End Sub



Private Sub CommandBars_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
        Dim Control As CommandBarControl, ControlItem As CommandBarControl
        
        If TypeOf CommandBar Is RibbonBackstageView Then
            'Debug.Print "RibbonBackstageView"
        End If
        
        Set Control = CommandBar.FindControl(, IDS_ARRANGE_BY)
        If Not Control Is Nothing Then
            Dim Index As Long
            Index = Control.Index
            Control.Visible = False
            
            Do While Index + 1 <= CommandBar.Controls.Count
                Set ControlItem = CommandBar.Controls.Item(Index + 1)
                If ControlItem.id = IDS_ARRANGE_BY Then
                    ControlItem.Delete
                Else
                    Exit Do
                End If
            Loop
            
'            Dim CurrentColumn As ReportColumn
'            For Each CurrentColumn In frmInbox. wndReportControl.Columns
'                Set ControlItem = CommandBar.Controls.Add(xtpControlButton, ID_REPORTCONTROL_COLUMN_ARRANGE_BY, CurrentColumn.Caption)
'                ControlItem.Parameter = CurrentColumn.ItemIndex
'                If Not frmInbox. wndReportControl.SortOrder.IndexOf(CurrentColumn) = -1 Then
'                    ControlItem.Checked = True
'                End If
'                If Not CurrentColumn.Visible Then
'                    ControlItem.Visible = False
'                End If
'            Next
        
        End If
End Sub

Private Sub CommandBars_SpecialColorChanged()
    Me.BackColor = CommandBars.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Error Resume Next
    'Debug.Print Control.Name
    
    Select Case Control.id
        Case ID_VIEW_STATUSBAR:     Control.Checked = CommandBars.statusBar.Visible
        
        
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            Control.Checked = IIf(frmInbox.wndCalendarControl.ViewType = xtpCalendarWorkWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_WEEK:
            Control.Checked = IIf(frmInbox.wndCalendarControl.ViewType = xtpCalendarWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_MONTH:
            Control.Checked = IIf(frmInbox.wndCalendarControl.ViewType = xtpCalendarMonthView, True, False)
        
        Case ID_OPTIONS_ANIMATION:
            Control.Checked = CommandBars.ActiveMenuBar.EnableAnimation
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
                Dim newFontHeight As Integer
                newFontHeight = FontSizes(Control.id - ID_OPTIONS_FONT_SYSTEM)
                Control.Checked = IIf(RibbonBar.FontHeight = newFontHeight, True, False)
                
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
                Control.Checked = CommandBars.PaintManager.AutoResizeIcons

        Case ID_RIBBON_EXPAND:
            Control.Visible = RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            Control.Visible = Not RibbonBar.Minimized
    End Select
   
End Sub

Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    If (Action = PaneActionSplitterResized) Then
        DockingPaneManager.RecalcLayout
        
        ' Save MRUShortcutBarWidth
        If (frmShortBar.ScaleWidth > MinimizedShortcutBarWidth And Container.Container.Type = PaneTypeSplitterContainer) Then
            'Debug.Print frmShortBar.ScaleWidth
            MRUShortcutBarWidth = frmShortBar.ScaleWidth
        End If
    Else
        If (Action = PaneActionSplitterResized) Then Debug.Print "Resizing "
    End If
End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Tag = PANE_SHORTCUTBAR Then
        Item.Handle = frmShortBar.hwnd
    ElseIf Item.Tag = PANE_REPORT_CONTROL Then
        Item.Handle = frmInbox.hwnd
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaDatosMenusDemas()
    
    CreateRibbon
    CreateBackstage
    CreateRibbonOptions
    CargaMenu
    CreateStatusBar
    PonerCaption
    CreateCalendarTabOriginal
    RibbonSeHaCreado = True
End Sub





Public Sub CambiarEmpresa(QueEmpresa As Integer)

    CadenaDesdeOtroForm = vUsu.Login & "|" & vEmpresa.codempre & "|"
        
    
        
    Set vUsu = New Usuario
    vUsu.Leer RecuperaValor(CadenaDesdeOtroForm, 1)
    
    vUsu.CadenaConexion = "arigestion" & QueEmpresa
    
    
    AbrirConexion vUsu.CadenaConexion
    
    Set vEmpresa = New Cempresa
    Set vParam = New Cparametros
    
    'NO DEBERIAN DAR ERROR
    vEmpresa.Leer
    vParam.Leer

    
    PonerCaption
    
   CargaDatosMenusDemas
   frmPaneContacts.SeleccionarNodoEmpresa vEmpresa.codempre
   pageBackstageHelp.Label9.Caption = vEmpresa.nomempre
   pageBackstageHelp.tabPage(0).Visible = False
   pageBackstageHelp.tabPage(1).Visible = False
   Me.RibbonBar.RedrawBar
   
    vControl.UltEmpre = vUsu.CadenaConexion
    vControl.Grabar
    
End Sub



Private Sub Form_Load()
   
    'Cargamos librerias de icinos de los forms
    frmIdentifica.pLabel "Carga DLL"
    CargaIconosDlls
   
    CommandBarsGlobalSettings.App = App
            
    frmIdentifica.pLabel "Leyendo menus usuario"
    CargaDatosMenusDemas
    
    ShowEventInPane = False
       
    FontSizes(0) = 0
    FontSizes(1) = 11
    FontSizes(2) = 13
    FontSizes(3) = 16
               
    DockingPaneManager.SetCommandBars Me.CommandBars
              
    Set frmShortBar = New frmShortcutBar2
    Set frmInbox = New frmInbox
        
    Dim A As Pane, B As Pane, C As Pane, d As Pane
    
    frmIdentifica.pLabel "Creando paneles"
    Set A = DockingPaneManager.CreatePane(PANE_SHORTCUTBAR, 170, 120, DockLeftOf, Nothing)
    A.Tag = PANE_SHORTCUTBAR
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    
    Set B = DockingPaneManager.CreatePane(PANE_REPORT_CONTROL, 700, 400, DockRightOf, A)
    B.Tag = PANE_REPORT_CONTROL
   
    DockingPaneManager.Options.HideClient = True
        
    Set CommandBars.Icons = CommandBarsGlobalSettings.Icons
    LoadIcons
    
    DockingPaneManager.RecalcLayout
    MRUShortcutBarWidth = frmShortBar.ScaleWidth
   
   
    'En funcion
    ' ID_OPTIONS_STYLEBLUE2010  ID_OPTIONS_STYLESILVER2010    ID_OPTIONS_STYLEBLACK2010
    frmIdentifica.pLabel "Carga skin"
    If vUsu.Skin = 3 Then
        Cad = ID_OPTIONS_STYLEBLACK2010
    Else
        If vUsu.Skin = 2 Then
            Cad = ID_OPTIONS_STYLESILVER2010
        Else
            Cad = ID_OPTIONS_STYLEBLUE2010
        End If
    End If
    CommandBars.FindControl(, Cad, , True).Execute

    'Por si se hubiera quedado bloqueado algo
    BorrarZBloqueos
End Sub


Private Sub CargaIconosDlls()

    ImageList1.ImageHeight = 48
    ImageList1.ImageWidth = 48
    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 48


    ImageList2.ImageHeight = 16
    ImageList2.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 16

    ImageListPPal48.ImageHeight = 48
    ImageListPPal48.ImageWidth = 48
    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 8, 48


    ImageListPpal16.ImageHeight = 16
    ImageListPpal16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 9, 16

'    Me.Icon = Me.ImageListPpal16.ListImages(2).Picture


    ImgListComun.ImageHeight = 24
    ImgListComun.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 2, 24 'antes icolistcon
    
    '++
    imgListComun_BN.ImageHeight = 24
    imgListComun_BN.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 3, 24
    
    imgListComun_OM.ImageHeight = 24
    imgListComun_OM.ImageWidth = 24
    GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 4, 24
    
    imgListComun16.ImageHeight = 16
    imgListComun16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 5, 16
    
    GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 6, 16
    GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 7, 16


End Sub

Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal op As Integer, ByVal tam As Integer)
    Dim I As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub



Public Sub ExpandButtonClicked()
   
    
    
    Dim A As Pane
    Set A = DockingPaneManager.FindPane(PANE_SHORTCUTBAR)
    
    Dim ShortcutBarMinimized As Boolean
    ShortcutBarMinimized = frmShortBar.ScaleWidth <= MinimizedShortcutBarWidth
    
    Dim NewWidth As Long
    If (ShortcutBarMinimized) Then
        NewWidth = MRUShortcutBarWidth
    Else
        NewWidth = MinimizedShortcutBarWidth
        frmShortBar.wndShortcutBar.PopupWidth = MRUShortcutBarWidth
    End If
        
    
    ' Set Size of Pane
    A.MinTrackSize.Width = NewWidth
    A.MaxTrackSize.Width = NewWidth
        
    DockingPaneManager.RecalcLayout
    DockingPaneManager.NormalizeSplitters
    DockingPaneManager.RedrawPanes
    
    ' Restore Constraints
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    A.MaxTrackSize.Width = 32000
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (pageBackstageInfo Is Nothing) Then Unload pageBackstageInfo
    If Not (pageBackstageHelp Is Nothing) Then Unload pageBackstageHelp
    If Not (pageBackstageSend Is Nothing) Then Unload pageBackstageSend
    
    'close all sub forms
    On Error Resume Next
    Dim I As Long
    For I = Forms.Count - 1 To 1 Step -1
        
        Unload Forms(I)
    Next
End Sub

Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, id, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddButton = Control
    
End Function

Private Sub CommandBars_Resize()
    
    On Error Resume Next
    
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.GetClientRect Left, top, Right, Bottom
    
End Sub

Private Sub LoadIcons()
    CommandBars.Icons.RemoveAll
    SuiteControlsGlobalSettings.Icons.RemoveAll
    ReportControlGlobalSettings.Icons.RemoveAll

    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\help.png", ID_APP_ABOUT, xtpImageNormal
        
        
        
   
    'Para que no carge imagen de ratios y graficas y punteo, no lo pongo aqui ya que los cargo "pequeños"
    '
  
      
    'ICONOS PEQUEÑOS
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(1, 1, 1, 1, 1, 1), xtpImageNormal
        
     
    
    'Pequeños
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1), xtpImageNormal
        
    'Pequeños diario
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(1, 1, 1, 1, 1, 1), xtpImageNormal
      
   
    'Deberiamos cargar un array con unos(1) de longitud 143
    ' y en funcion del valor del campo imagen en el punto de menu correspondiente
    ' lo pondremos en el array.
    ' Ejemplo    303 Extractos  Campo imagen: 87
    ' quiere decir que en el campo 87 del array sustituieremos el 1 por el 303


'
    Dim T() As Variant
    'Cad linea son 15
    T = Array(1, 1, ID_Empresa, 1, 1, ID_Parámetros, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, ID_Contadores, 1, 1, 1, 1, 1, 1, 1, 1, _
        ID_Informes, 1, ID_Clientes, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, ID_Usuarios, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, ID_ConceptosFacturas, ID_Expedientes, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, ID_PagodeTasas, 1, ID_Caja, 1, 1, ID_FacturasEmitidas, 1, 1, ID_PrevisionFacturacion, 1, 1, 1, 1, _
        1, 1, 1, ID_PrevisionFacturacion, 1, ID_IntegraciónContable, 1, 1, ID_EstadisticaClientes, 1, 1, 1, 1, 1, 1, _
        1, 1, ID_EstadisticaConceptos, 1, 1, ID_Gráficamensual, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1)
    
     'ID_Empresa , ID_Parámetros, ID_Contadores , ID_Usuarios , ID_Informes , ID_Clientes , ID_ConceptosFacturas ,
     'ID_Expedientes , ID_PagodeTasas , ID_Caja , ID_FacturasEmitidas, ID_PrevisionFacturacion, ID_Facturasdirectas,
     'ID_Facturasperiodicas , ID_IntegraciónContable, ID_EstadisticaClientes, ID_EstadisticaConceptos, ID_Gráficamensual
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", T, xtpImageNormal
    
           

    'Este de abjo funciona correctamente.
    'NO tocar. Es por si falla volver a empezar
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", _
'            Array(ID_CarteradeCobros, ID_InformeCobrosPendientes, ID_RealizarCobro, ID_Compensarcliente, 1, ID_BalancePresupuestario, 1, _
'            ID_CentrosdeCoste, 1, 1, ID_Presupuestos, ID_Remesas, ID_Detalledeexplotación, ID_CarteradePagos, ID_CuentadeExplotaciónAnalítica, ID_ExtractosporCentrodeCoste, _
'            ID_Asientos, ID_Extractos, ID_Punteo, 1, ID_CuentadeExplotación, ID_Totalesporconcepto, ID_BalancedeSituación, ID_PérdidasyGanancias, _
'            ID_SumasySaldos, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            ID_Empresa, ID_ParametrosContabilidad, ID_Contadores, ID_Usuarios, 1, ID_Informes, ID_Nuevaempresa, ID_ConfigurarBalances, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            ID_FacturasEmitidas, ID_LibroFacturasEmitidas, ID_FacturasRecibidas, ID_LibroFacturasRecibidas, 1, 1, 1, 1, 1, ID_Elementos, ID_GenerarAmortización, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            1, ID_PlanContable, ID_TiposdeDiario, ID_Conceptos, ID_TiposdeIVA, ID_TiposdePago, ID_Bancos, ID_FormasdePago, _
'            ID_BicSwift, ID_Agentes, ID_AsientosPredefinidos, ID_ModelosdeCartas, _
'            ID_Renumeracióndeasientos, ID_CierredeEjercicio, ID_Deshacercierre, 1, 1, 1, 1, 1, 1, ID_DiarioOficial, _
'            ID_PresentaciónTelemáticadeLibros, ID_Traspasodecuentasenapuntes, ID_Renumerarregistrosproveedor, 1, ID_TraspasocodigosdeIVA), xtpImageNormal
'
    
    'Presupuiestaria y analitaica cargadas arriba en pequeño
    '---------------------------------------------------------
    '
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(1, 1, 1, 1, 1, 1), xtpImageNormal
    

    

    'Pequeños
    ' ID_Compensaciones ID_Reclamaciones  ID_InformeImpagados ID_RemesasTalenPagare ID_Norma57Pagosventanilla  ID_TransferenciasAbonos
    ' ID_InformePagosbancos ID_Transferencias ID_Pagosdomiciliados ID_GastosFijos ID_Compensarproveedor ID_Confirming
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(1, 1, 1, 1, 1, 1, 1, _
            1, 1, 1), xtpImageNormal
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(1, 1, 1, 1, 1, 1), xtpImageNormal
    
     
 
        
        
    '------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookcalicons.png", _
            Array(ID_GROUP_NEW_APPOINTMENT, ID_GROUP_NEW_MEETING, ID_GROUP_NEW_ITEMS, ID_GROUP_GOTO_TODAY, _
            ID_GROUP_GOTO_NEXT7DAYS, ID_GROUP_ARRANGE_DAY, ID_GROUP_ARRANGE_WORK_WEEK, ID_GROUP_ARRANGE_WEEK, _
            ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_SCHEDULE_VIEW, ID_GROUP_MANAGE_CALENDARS_OPEN, ID_GROUP_MANAGE_CALENDARS_GROUPS, _
            ID_GROUP_SHARE_EMAIL, ID_GROUP_SHARE_SHARE, ID_GROUP_SHARE_PUBLISH, ID_GROUP_SHARE_PERMISSIONS), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\RibbonMinimize.png", _
            Array(ID_RIBBON_MINIMIZE, ID_RIBBON_EXPAND), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\Search.png", _
            ID_SEARCH_ICON, xtpImageNormal
            
     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonslarge.png", _
            Array(ID_GROUP_MAIL_NEW_NEW, ID_GROUP_MAIL_NEW_NEW_ITEMS, ID_GROUP_MAIL_DELETE_DELETE, ID_GROUP_MAIL_RESPOND_REPLY, _
            ID_GROUP_MAIL_RESPOND_REPLY_ALL, ID_GROUP_MAIL_RESPOND_FORWARD, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
            
     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", _
            Array(ID_GROUP_MAIL_DELETE_CLEANUP, ID_GROUP_MAIL_DELETE_JUNK, ID_GROUP_MAIL_RESPOND_MEETING, ID_GROUP_MAIL_RESPOND_IM, _
            ID_GROUP_MAIL_RESPOND_MORE, ID_GROUP_MAIL_TAGS_UNREAD, ID_GROUP_MAIL_TAGS_CATEGORIZE, ID_GROUP_MAIL_TAGS_FOLLOWUP, ID_GROUP_MAIL_FIND_ADDRESSBOOK, _
            ID_GROUP_MAIL_FIND_FILTER, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
    
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookpane.png", _
            Array(ID_SWITCH_NORMAL, ID_SWITCH_CALENAR_AND_TASK, ID_SWITCH_CALENDAR, ID_SWITCH_CLASSIC, ID_SWITCH_READING), xtpImageNormal
            
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_24x24.bmp", _
            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
            
        CommandBars.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_QUICKSTEP_REPLAY_DELETE, ID_QUICKSTEP_TO_MANAGER, ID_QUICKSTEP_MOVE_TO, ID_QUICKSTEP_CREATE_NEW, ID_QUICKSTEP_TEAM_EMAIL, ID_QUICKSTEP_DONE), xtpImageNormal
            
        ReportControlGlobalSettings.Icons.LoadBitmap App.Path & "\styles\bmreport.bmp", _
        Array(COLUMN_MAIL_ICON, COLUMN_IMPORTANCE_ICON, COLUMN_CHECK_ICON, RECORD_UNREAD_MAIL_ICON, RECORD_READ_MAIL_ICON, _
            RECORD_REPLIED_ICON, RECORD_IMPORTANCE_HIGH_ICON, COLUMN_ATTACHMENT_ICON, COLUMN_ATTACHMENT_NORMAL_ICON, _
            RECORD_IMPORTANCE_LOW_ICON), xtpImageNormal
            
        Dim I As Integer
        For I = 1 To 17
            SuiteControlsGlobalSettings.Icons.LoadIcon App.Path & "\styles\TreeView\icon" & I & ".ico", I, xtpImageNormal
        Next I
End Sub

Private Sub SaveRibbonBarToXML()
    Dim Px As PropExchange
    Set Px = XtremeCommandBars.CreatePropExchange()
    
    Px.CreateAsXML False, "Settings"
        
    Dim Options As StateOptions
    Set Options = CommandBars.CreateStateOptions()
    Options.SerializeControls = True
        
    CommandBars.DoPropExchange Px.GetSection("CommandBars"), Options
    
    Px.SaveToFile "C:\Layout.xml"
    
End Sub



Private Function CreateQuickStepGallery() As CommandBarGalleryItems

    Dim GalleryItems As CommandBarGalleryItems
    Set GalleryItems = CommandBars.CreateGalleryItems(ID_GALLERY_QUICKSTEP)
        
    GalleryItems.ItemWidth = 120
    GalleryItems.ItemHeight = 20
            
    GalleryItems.AddItem ID_QUICKSTEP_MOVE_TO, "Move To: ?"
    GalleryItems.AddItem ID_QUICKSTEP_TO_MANAGER, "To Manager"
    GalleryItems.AddItem ID_QUICKSTEP_TEAM_EMAIL, "Team E-mail"
    GalleryItems.AddItem ID_QUICKSTEP_DONE, "Done"
    GalleryItems.AddItem ID_QUICKSTEP_REPLAY_DELETE, "Reply & Delete"
    GalleryItems.AddItem ID_QUICKSTEP_CREATE_NEW, "Create New"
        
    GalleryItems.Icons = CommandBarsGlobalSettings.Icons

    Set CreateQuickStepGallery = GalleryItems

End Function

Private Sub CommandBars_ControlNotify(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal Code As Long, ByVal NotifyData As Variant, Handled As Variant)
    If (Code = XTP_BS_TABCHANGED) Then

        
    End If
End Sub


Private Sub CreateBackstage()

    
    Dim RibbonBar As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
    Dim BackstageView As RibbonBackstageView
    Set BackstageView = CommandBars.CreateCommandBar("CXTPRibbonBackstageView")
    
    BackstageView.SetTheme xtpThemeRibbon


    CommandBars.Icons.LoadBitmap App.Path & "\styles\BackstageIcons.png", _
    Array(1, 1, 1002, 1, 1, ID_APP_EXIT), xtpImageNormal

    Set RibbonBar.AddSystemButton.CommandBar = BackstageView
    
    'BackstageView.AddCommand ID_FILE_SAVE, "Cambiar empresa"
    'BackstageView.AddCommand ID_FILE_SAVE_AS, "Personalizar"
    'BackstageView.AddCommand ID_FILE_OPEN, "Open"
    'BackstageView.AddCommand ID_FILE_CLOSE, "Close"
    
    If (pageBackstageInfo Is Nothing) Then Set pageBackstageInfo = New pageBackstageInfo
    If (pageBackstageSend Is Nothing) Then Set pageBackstageSend = New pageBackstageSend
    If (pageBackstageHelp Is Nothing) Then Set pageBackstageHelp = New pageBackstageHelp
    
    Dim ControlInfo As RibbonBackstageTab
    Set ControlInfo = BackstageView.AddTab(1000, "Info", pageBackstageHelp.hwnd)
    
    BackstageView.AddTab 1002, "Empresas", pageBackstageSend.hwnd

    ' Los menus de informacion...
    BackstageView.AddTab 1001, "Acerca de", pageBackstageInfo.hwnd
    
    
    
    
    
    
    
    
    
    
    'BackstageView.AddCommand ID_FILE_OPTIONS, "Options"
    BackstageView.AddCommand ID_APP_EXIT, "Salir"
    
    ControlInfo.DefaultItem = True
    

End Sub




Private Sub CreateCalendarTabOriginal()

    Dim TabCalendarHome As RibbonTab
    Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup

    
    Dim Control As CommandBarControl
    Dim ControlNew_NewItems As CommandBarPopup
    Dim ControlArrange_Month As CommandBarPopup
    Dim ControlManage_Open As CommandBarPopup
    Dim ControlManage_Groups As CommandBarPopup
    Dim ControlShare_Publish As CommandBarPopup
           
    Dim PopupBar As CommandBar
    
    Set TabCalendarHome = RibbonBar.InsertTab(14, "Agenda")
    TabCalendarHome.id = ID_TAB_CALENDAR_HOME
 
    Set GroupNew = TabCalendarHome.Groups.AddGroup("&Nueva", ID_GROUP_NEW)
        
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Evento")
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_MEETING, "&Cita")
    
    '------------------------------------
    'Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "Evento")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "E&vento todo el dia")
    '    Control.BeginGroup = True
    'ControlNew_NewItems.KeyboardTip = "V"
    
    Set GroupGoTo = TabCalendarHome.Groups.AddGroup("I&r a", ID_GROUP_GOTO)
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_TODAY, "&Hoy")
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_NEXT7DAYS, "Próximos &7 dias ")
    GroupGoTo.ShowOptionButton = True
    GroupGoTo.ControlGroupOption.Caption = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.ToolTipText = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.DescriptionText = "Ir a fecha especificada."
    
    Set GroupArrange = TabCalendarHome.Groups.AddGroup("Vista", ID_GROUP_ARRANGE2)
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_DAY, "&Dia vista")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WORK_WEEK, "Samana &trabajo")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WEEK, "Sema&na vista")
    Set ControlArrange_Month = GroupArrange.Add(xtpControlSplitButtonPopup, ID_GROUP_ARRANGE_MONTH, "Mes")
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_LOW, "Ver detalle")
            Control.ToolTipText = "Muestra solo eventos todo el dia."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_MEDIUM, "Detalle &Medio")
            Control.ToolTipText = "Eventos todo el dia y si esta libre el dia o tiene eventos."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_HIGH, "Detalle &Alto")
            Control.ToolTipText = "Muestra todo."
            Control.DescriptionText = Control.ToolTipText

'    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_SCHEDULE_VIEW, "Schedule View")
'    GroupArrange.ShowOptionButton = True
'    GroupArrange.ControlGroupOption.Caption = "Calendar Options"
'    GroupArrange.ControlGroupOption.ToolTipText = "Calendar Options"
'    GroupArrange.ControlGroupOption.DescriptionText = "Change the settings for calendars, meetings and time zones."
'
'
  
    
End Sub





Private Sub CreateRibbon()

    
    If RibbonSeHaCreado Then Exit Sub
    Dim RibbonBar As RibbonBar
    
    
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    
    RibbonBar.AllowQuickAccessCustomization = False
    RibbonBar.ShowQuickAccessBelowRibbon = False
    RibbonBar.ShowGripper = False
    
    RibbonBar.AllowMinimize = False
    RibbonBar.AddSystemButton
    
    RibbonBar.SystemButton.IconId = ID_SYSTEM_ICON
    RibbonBar.SystemButton.Caption = "&Menu"
    RibbonBar.SystemButton.Style = xtpButtonCaption
End Sub

Private Sub CreateRibbonOptions()

    CommandBars.EnableActions
    If RibbonSeHaCreado Then Exit Sub
    
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue", "Office 2010 Blue", "Office 2010 Blue", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver", "Office 2010 Silver", "Office 2010 Silver", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black", "Office 2010 Black", "Office 2010 Black", "Themes"

    Dim Control As CommandBarControl, ControlAbout As CommandBarControl
    Dim ControlPopup As CommandBarPopup, ControlOptions As CommandBarPopup
         
    Set ControlOptions = RibbonBar.Controls.Add(xtpControlPopup, 0, "Opciones")
    ControlOptions.Flags = xtpFlagRightAlign
    
    Set Control = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Styles")
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black"
    
    Set ControlPopup = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Tamaño fuente", -1, False)
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_SYSTEM, "Sistema", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlRadioButton, ID_OPTIONS_FONT_NORMAL, "Normal", -1, False)
    Control.BeginGroup = True
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_LARGE, "Grande", -1, False
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_EXTRALARGE, "Extra grande", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_OPTIONS_FONT_AUTORESIZEICONS, "Ajustar Icons", -1, False)
    Control.BeginGroup = True
    
    'ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_RTL, "Right To Left"
    ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_ANIMATION, "Animation   "
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_MINIMIZE, "Minimizar la barra", False, "Muestra solo los titulos del menu principal.")
    Control.Flags = xtpFlagRightAlign
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_EXPAND, "Expandir la barra", False, "Muestra todos los elementos del menu.")
    Control.Flags = xtpFlagRightAlign
        
    Set ControlAbout = RibbonBar.Controls.Add(xtpControlButton, ID_APP_ABOUT, "&Acerca de")
    ControlAbout.Flags = xtpFlagRightAlign Or xtpFlagManualUpdate
    

        
End Sub








'*************************************************************************
'*************************************************************************
'*************************************************************************
'
'       CARGA menus en Ribbon
'
'




Public Sub CargaMenu()
Dim RN As ADODB.Recordset




    Set RN = New ADODB.Recordset
    Set Rn2 = New ADODB.Recordset
    On Error GoTo eCargaMenu
        
    idTabPpal = 0
    If RibbonSeHaCreado Then RibbonBar.RemoveAllTabs
    
    Cad = "Select * from menus where aplicacion = 'arigestion' and padre =0 ORDER BY padre,orden "
    RN.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
    
        
        If Not BloqueaPuntoMenu(RN!Codigo, "arigestion") Then
             Habilitado = True
             
             If Not MenuVisibleUsuario(DBLet(RN!Codigo), "arigestion") Then
                 Habilitado = False
             Else
         
                 If (MenuVisibleUsuario(DBLet(RN!Padre), "arigestion") And DBLet(RN!Padre) <> 0) Or DBLet(RN!Padre) = 0 Then
                     'OK todo habilitado
                 Else
                     Habilitado = False
                 End If
             End If
             
      
                
            If Habilitado Then
                
                Select Case RN!Codigo
                Case 1
                    '1   "CONFIGURACION"
                    CargaMenuConfiguracion RN!Codigo
                    
                    
                ' ****  Iran todos juntos en un tab
                Case 2
                    '2 Datos generales
                    CargaMenuDatosGenerales RN!Codigo
                Case 3
                    '3   "TRABAJO DIARIO"
                    CargaMenuTrabajoDiario RN!Codigo
                Case 4
                    '4   "FACTURACION"
                    CargaMenuFacturacion RN!Codigo
                Case 5
                    '5   "ESTADISTICAS"
                    CargaMenuEstadistica RN!Codigo


                Case Else
                    MsgBox "Menu no tratado"
                    End
                End Select
                
            End If
                                                 
        End If  'de habilitado el padre
    
        RN.MoveNext
    Wend
    RN.Close
                        
               
    
        RibbonBar.Tab(idTabPpal).Visible = True
        RibbonBar.Tab(idTabPpal).Selected = True
        Set RibbonBar.SelectedTab = RibbonBar.Tab(idTabPpal)
      
    
    
eCargaMenu:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    Set TabNuevo = Nothing
    Set GroupNew = Nothing
    Set Control = Nothing
    Set RN = Nothing
    Set Rn2 = Nothing
End Sub



Private Sub CargaMenuConfiguracion(IdMenu As Integer)

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Configuracion")
        TabNuevo.id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
       
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "arigestion") Then Habilitado = False
                End If
           
           
                    
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close

         Set GroupNew = Nothing
End Sub



Private Sub CrearTabPPal()
    
    If idTabPpal = 0 Then
        Set TabNuevo = RibbonBar.InsertTab(9999, "Diario")
        idTabPpal = TabNuevo.Index
    Else
        Set TabNuevo = RibbonBar.Tab(idTabPpal)
    End If
End Sub


Private Sub CargaMenuDatosGenerales(IdMenu As Integer)

        'Creamos la TAB
        CrearTabPPal
        
        'En este llevaremos dos solapas, tesoreria y contabilidad (no le ponemos nombres)
        Cad = CStr(IdMenu * 100000)
        
        Set GroupNew = TabNuevo.Groups.AddGroup("General", Cad & "0")
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "arigestion") Then Habilitado = False
                End If
           
           
                    
              
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                
                 
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close

         Set GroupNew = Nothing
End Sub


Private Sub CargaMenuFacturacion(IdMenu As Integer)


        'Creamos la TAB
        CrearTabPPal
        
        
        Cad = CStr(IdMenu * 100000)
        Set GroupNew = TabNuevo.Groups.AddGroup("Facturación", Cad & "0")
        
    
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "arigestion") Then Habilitado = False
                End If
                

                
                
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                
                Control.Enabled = Habilitado
                
              
              
              
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub


Private Sub CargaMenuEstadistica(IdMenu As Integer)
'Dim GropCli As RibbonGroup
'Dim GrupPag As RibbonGroup
        

        
        'Creamos la TAB
        CrearTabPPal
        
        Cad = CStr(IdMenu * 100000)
        Set GroupNew = TabNuevo.Groups.AddGroup("Estadística", Cad & "2")
    


'
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "arigestion") Then Habilitado = False
                End If
            End If
            
            
'            Select Case Rn2!Codigo
'            Case 401, 402, 403
'                Set Control = GropCli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
 '           Case 404, 405, 406
 '               Set Control = GrupPag.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
 '           Case Else
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
 '           End Select
            
            
            Cad = "NO"
            Control.Enabled = Habilitado
           ' ControlNew_NewItems.KeyboardTip = "V"
         
            Rn2.MoveNext
        Wend
        Rn2.Close


End Sub








Private Sub CargaMenuTrabajoDiario(IdMenu As Integer)
Dim Col As Collection

        
        
        
        'Este veremos si tiene alguna utilidad activa. Si es asi, crearemos la solapa, si no nada
        '.......................................................................
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'arigestion' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        Set Col = New Collection
        While Not Rn2.EOF
           I = I + 1
           If Not BloqueaPuntoMenu(Rn2!Codigo, "arigestion") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "arigestion") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "arigestion") Then Habilitado = False
                End If
            End If
            
            Col.Add Abs(Habilitado) & "|" & Rn2!Codigo & "|" & Rn2!Descripcion & "|"
            If Habilitado Then Cad = "S"
            
            Rn2.MoveNext
        Wend
        Rn2.Close
        
            '1408    "Traspaso de cuentas en apuntes"
            '1409    "Renumerar registros proveedor"
            '1410    "Aumentar dígitos contables"
            '1411    "Traspaso códigos de I.V.A."
            '1412    "Acciones realizadas"
            '1413    Importar fras cliente
            
        'Ya puedo utilizar numregelim
        If Cad <> "" Then
            'OK creamos solapa y demas
            'Creamos la TAB
            'Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Trabajo diario")
            'TabNuevo.id = CLng(IdMenu)
            CrearTabPPal
            Set GroupNew = TabNuevo.Groups.AddGroup("Trabajo diario", 14000001)
            For NumRegElim = 1 To Col.Count
                Habilitado = CStr(RecuperaValor(Col.Item(NumRegElim), 1)) = "1"
                Set Control = GroupNew.Add(xtpControlButton, CLng(RecuperaValor(Col.Item(NumRegElim), 2)), CStr(RecuperaValor(Col.Item(NumRegElim), 3)))
                Control.Enabled = Habilitado
            Next
                
            
        End If
        

Set Col = Nothing
End Sub






'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
Private Sub AbrirFormularios(Accion As Long)


'''''''Public Const ID_Empresa = 101
'''''''Public Const ID_Parámetros = 102
'''''''Public Const ID_Contadores = 103
'''''''Public Const ID_Usuarios = 104
'''''''Public Const ID_Informes = 105
'''''''Public Const ID_Clientes = 201
'''''''Public Const ID_ConceptosFacturas = 202
'''''''Public Const ID_Expedientes = 301
'''''''Public Const ID_PagodeTasas = 302
'''''''Public Const ID_Caja = 303
'''''''Public Const ID_FacturasEmitidas = 401
'''''''Public Const ID_PrevisionFacturacion = 402
'''''''Public Const ID_Facturasdirectas = 403
'''''''Public Const ID_Facturasperiodicas = 404
'''''''Public Const ID_IntegraciónContable = 405
'''''''Public Const ID_EstadisticaClientes = 501
'''''''Public Const ID_EstadisticaConceptos = 502
'''''''Public Const ID_Gráficamensual = 503
'''''''
'''''''Public Const ID_Licencia_Usuario_Final_txt = 2001
'''''''Public Const ID_Licencia_Usuario_Final_web = 2002
'''''''Public Const ID_Ver_Version_operativa_web = 2003



    Select Case Accion
    Case ID_Empresa
        frmempresa.Show vbModal
    Case ID_Parámetros
        frmClienteAcciones.Show vbModal
    Case ID_Contadores
        frmContadores.Show vbModal
    
    Case ID_Informes
        frmCrystal.Show vbModal
    Case ID_Usuarios
        frmMantenusu.Show vbModal
    Case ID_ConceptosFacturas
        frmConceptos.Show vbModal
    
    Case ID_Clientes
        'Load frmcolClientes
        'frmcolClientes.SetColor Id
        frmcolClientes.Show vbModal
    Case ID_Expedientes
        frmExpediente.numExpediente = ""
        frmExpediente.Show vbModal
        
    Case ID_FacturasEmitidas
        frmFacturasCli.FACTURA = ""
        frmFacturasCli.Show vbModal
        
    Case ID_PrevisionFacturacion
        frmPrevisionFacturacion.Show vbModal
    End Select

End Sub






'Esto lo tiene Moni "asin", ni digo ni pregunto
Private Sub AbrirListado(numero As Byte, Cerrado As Boolean)
'    Screen.MousePointer = vbHourglass
'    frmListado.EjerciciosCerrados = Cerrado
'    frmListado.Opcion = numero
'    frmListado.Show vbModal
End Sub




'Establecer y fijar Skin
Public Sub EstablecerSkin(QueSkin As Integer)

    FijaSkin QueSkin

  ' Cargando el archivo del Skin
  ' ============================
    'frmPpal.SkinFramework1.LoadSkin Skn$, ""
    Me.SkinFramework1.ApplyWindow frmppal.hwnd
    Me.SkinFramework1.ApplyOptions = Me.SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
    
  

    
End Sub

Private Function FijaSkin(numero)


  Select Case (numero)
 
           
            Case 1:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlue.ini"
            Case 2:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalSilver.ini"
            Case 3:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlack.ini"
                
                  
                
        
        
  End Select
    
End Function



Private Sub PonerCaption()
   '     Caption = "AriCONTA 6    V-" & App.Major & "." & App.Minor & "." & App.Revision & "    usuario: " & vUsu.Nombre & "      Ejercicio: " & vParam.fechaini & " - " & vParam.fechafin
        'Label33.Caption = "   " & vEmpresa.nomempre
End Sub


Public Sub OpcionesMenuInformacion(id As Long)
    
    Select Case id
    Case ID_Licencia_Usuario_Final_txt
        LanzaVisorMimeDocumento Me.hwnd, "c:\programas\Ariadna.rtf"
    Case ID_Licencia_Usuario_Final_web
        LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "AriCONTA-6.html?Licenciadeuso.html"
    Case ID_Ver_Version_operativa_web
        LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "Ariconta-6.html"  ' "http://www.ariadnasw.com/clientes/"
    End Select
    
End Sub

