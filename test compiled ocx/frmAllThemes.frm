VERSION 5.00
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#7.0#0"; "NewTab01.ocx"
Begin VB.Form frmAllThemes 
   Caption         =   "All current default themes"
   ClientHeight    =   11748
   ClientLeft      =   276
   ClientTop       =   840
   ClientWidth     =   22140
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11748
   ScaleWidth      =   22140
   StartUpPosition =   2  'CenterScreen
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   643
      Themed          =   0   'False
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(0)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   1
      Left            =   3350
      TabIndex        =   1
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      Style           =   0
      TabHeight       =   643
      HighlightEffect =   0   'False
      Themed          =   0   'False
      SoftEdges       =   0   'False
      AutoTabHeight   =   -1  'True
      HighlightMode   =   1
      HighlightModeTabSel=   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(1)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   1
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   2
      Left            =   6460
      TabIndex        =   2
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ControlJustAdded=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      TabHeight       =   643
      Themed          =   0   'False
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(2)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   3
      Left            =   9570
      TabIndex        =   3
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      Style           =   1
      TabHeight       =   643
      HighlightEffect =   0   'False
      Themed          =   0   'False
      SoftEdges       =   0   'False
      AutoTabHeight   =   -1  'True
      HighlightMode   =   1
      HighlightModeTabSel=   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(3)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   4
      Left            =   12680
      TabIndex        =   4
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      TabHeight       =   643
      Themed          =   0   'False
      TabWidthStyle   =   1
      ShowRowsInPerspective=   1
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(4)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   5
      Left            =   15790
      TabIndex        =   5
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      Style           =   2
      TabHeight       =   643
      HighlightEffect =   0   'False
      Themed          =   0   'False
      SoftEdges       =   0   'False
      AutoTabHeight   =   -1  'True
      HighlightMode   =   1
      HighlightModeTabSel=   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(5)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   5
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   6
      Left            =   18900
      TabIndex        =   6
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ControlJustAdded=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      TabHeight       =   643
      Themed          =   0   'False
      TabWidthStyle   =   0
      AutoTabHeight   =   -1  'True
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(6)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   728
      Themed          =   0   'False
      BackColorTabs   =   15658734
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   14181684
      HighlightMode   =   2
      HighlightModeTabSel=   66
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(7)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   7
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   8
      Left            =   3350
      TabIndex        =   8
      Top             =   2280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColorHighlighted=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   728
      HighlightEffect =   0   'False
      Themed          =   0   'False
      BackColorTabs   =   14611960
      BackColorTabSel =   16383485
      FlatBarColorHighlight=   3431538
      FlatBarColorInactive=   13559786
      FlatBorderColor =   1148870
      HighlightColor  =   3431538
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   1148870
      IconColorMouseHover=   16777215
      IconColorTabHighlighted=   16777215
      HighlightMode   =   68
      HighlightModeTabSel=   90
      FlatBorderMode  =   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(8)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00F9FDFD&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   8
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   9
      Left            =   6460
      TabIndex        =   9
      Top             =   2280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColorHighlighted=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   728
      HighlightEffect =   0   'False
      Themed          =   0   'False
      BackColorTabs   =   15136990
      BackColorTabSel =   16514553
      FlatBarColorHighlight=   3633716
      FlatBarColorInactive=   14150350
      FlatBorderColor =   1820177
      HighlightColor  =   3633716
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   1820177
      IconColorMouseHover=   16777215
      IconColorTabHighlighted=   16777215
      HighlightMode   =   68
      HighlightModeTabSel=   90
      FlatBorderMode  =   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(9)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FBFDF9&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   9
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   10
      Left            =   9570
      TabIndex        =   10
      Top             =   2280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   728
      Themed          =   0   'False
      BackColorTabs   =   15202556
      BackColorTabSel =   16777215
      FlatBarColorHighlight=   3530228
      FlatBarColorInactive=   13559786
      HighlightColor  =   3530228
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   768981
      IconColorMouseHover=   12664841
      IconColorTabHighlighted=   12664841
      HighlightMode   =   76
      HighlightModeTabSel=   90
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(10)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   10
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   11
      Left            =   12680
      TabIndex        =   11
      Top             =   2280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColorTabSel =   10184001
      ForeColorHighlighted=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   643
      Themed          =   0   'False
      BackColorTabSel =   16250871
      FlatTabsSeparationLineColor=   -2147483633
      FlatBorderColor =   10184001
      HighlightColor  =   16477710
      AutoTabHeight   =   -1  'True
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   10184001
      IconColorMouseHover=   16777215
      IconColorMouseHoverTabSel=   10184001
      IconColorTabHighlighted=   16777215
      HighlightMode   =   4
      HighlightModeTabSel=   10
      FlatBorderMode  =   1
      FlatBarHeight   =   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(11)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00F7F7F7&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H009B6541&
         Height          =   750
         Index           =   11
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   12
      Left            =   15790
      TabIndex        =   12
      Top             =   2280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   728
      Themed          =   0   'False
      BackColorTabs   =   1588736
      BackColorTabSel =   2183936
      FlatBarColorHighlight=   9615225
      FlatBarColorInactive=   4422175
      FlatTabsSeparationLineColor=   1983492
      FlatBodySeparationLineColor=   1983492
      FlatBorderColor =   3960091
      HighlightColor  =   5281568
      HighlightColorTabSel=   5942308
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   16777215
      TabMousePointerHand=   -1  'True
      HighlightMode   =   64
      HighlightModeTabSel=   90
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(12)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00215300&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   12
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   13
      Left            =   18898
      TabIndex        =   13
      Top             =   2280
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   728
      Themed          =   0   'False
      BackColorTabs   =   1835059
      BackColorTabSel =   2424902
      FlatBarColorHighlight=   8542630
      FlatBarColorInactive=   4332134
      FlatTabsSeparationLineColor=   2032442
      FlatBodySeparationLineColor=   2032442
      FlatBorderColor =   3937884
      HighlightColor  =   5184382
      HighlightColorTabSel=   5906065
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   16777215
      TabMousePointerHand=   -1  'True
      HighlightMode   =   64
      HighlightModeTabSel=   90
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(13)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00250046&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   13
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   14
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ControlJustAdded=   0   'False
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   728
      Themed          =   0   'False
      BackColorTabs   =   5057536
      BackColorTabSel =   6699520
      FlatBarColorHighlight=   16777215
      FlatBarColorInactive=   9856549
      FlatTabsSeparationLineColor=   5386501
      FlatBodySeparationLineColor=   5057536
      FlatBorderColor =   8870689
      HighlightColor  =   7554571
      HighlightColorTabSel=   13863980
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   16777215
      TabMousePointerHand=   -1  'True
      HighlightMode   =   76
      HighlightModeTabSel=   90
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(14)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00663A00&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   14
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   15
      Left            =   3350
      TabIndex        =   15
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColorTabSel =   16731706
      ForeColorHighlighted=   16731706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   749
      Themed          =   0   'False
      BackColorTabs   =   16250871
      FlatBarColorHighlight=   16250871
      FlatBarColorInactive=   16250871
      FlatTabsSeparationLineColor=   16250871
      FlatBodySeparationLineColor=   14869218
      FlatBorderColor =   16250871
      HighlightColor  =   16477710
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   16731706
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   16731706
      IconColorMouseHover=   16731706
      IconColorMouseHoverTabSel=   16731706
      IconColorTabHighlighted=   16731706
      HighlightMode   =   1
      HighlightModeTabSel=   64
      FlatBarHeight   =   4
      FlatBarPosition =   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(15)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   15
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   16
      Left            =   6460
      TabIndex        =   16
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   0
      ForeColorTabSel =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   643
      Themed          =   0   'False
      BackColorTabSel =   16250871
      FlatTabsSeparationLineColor=   11250603
      FlatBodySeparationLineColor=   7699508
      FlatBorderColor =   7699508
      HighlightColor  =   12766860
      HighlightColorTabSel=   7699508
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7699508
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   16777215
      IconColorMouseHoverTabSel=   16777215
      HighlightMode   =   4
      HighlightModeTabSel=   20
      FlatBorderMode  =   1
      FlatBarHeight   =   0
      FlatBodySeparationLineHeight=   3
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(16)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00F7F7F7&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   16
         Left            =   120
         TabIndex        =   48
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   17
      Left            =   9570
      TabIndex        =   17
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   855
      Themed          =   0   'False
      BackColorTabs   =   12766860
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   7699508
      FlatBarColorInactive=   12766860
      FlatTabsSeparationLineColor=   11250603
      FlatBodySeparationLineColor=   11250603
      FlatBorderColor =   -2147483633
      HighlightColor  =   12766860
      HighlightColorTabSel=   -2147483633
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7699508
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      TabMousePointerHand=   -1  'True
      HighlightMode   =   196
      HighlightModeTabSel=   196
      FlatBarPosition =   1
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(17)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   17
         Left            =   120
         TabIndex        =   49
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   18
      Left            =   12680
      TabIndex        =   18
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   855
      Themed          =   0   'False
      BackColorTabs   =   12766860
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   7699508
      FlatBarColorInactive=   12766860
      FlatTabsSeparationLineColor=   11250603
      FlatBodySeparationLineColor=   11250603
      FlatBorderColor =   -2147483633
      HighlightColor  =   12766860
      HighlightColorTabSel=   -2147483633
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7699508
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      FlatRoundnessTabs=   16
      TabMousePointerHand=   -1  'True
      HighlightMode   =   196
      HighlightModeTabSel=   196
      FlatBarPosition =   1
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(18)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   18
         Left            =   120
         TabIndex        =   50
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   19
      Left            =   15788
      TabIndex        =   19
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   855
      Themed          =   0   'False
      BackColorTabs   =   12766860
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   7699508
      FlatBarColorInactive=   12766860
      FlatTabsSeparationLineColor=   11250603
      FlatBodySeparationLineColor=   7699508
      FlatBorderColor =   -2147483633
      HighlightColor  =   12766860
      HighlightColorTabSel=   -2147483633
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7699508
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      TabMousePointerHand=   -1  'True
      HighlightMode   =   196
      HighlightModeTabSel=   1
      FlatBodySeparationLineHeight=   3
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(19)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   19
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   20
      Left            =   18896
      TabIndex        =   20
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   897
      Themed          =   0   'False
      BackColorTabs   =   12766860
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   7699508
      FlatBarColorInactive=   12766860
      FlatTabsSeparationLineColor=   -2147483633
      FlatBodySeparationLineColor=   7699508
      FlatBorderColor =   7699508
      HighlightColor  =   12766860
      HighlightColorTabSel=   -2147483633
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7699508
      FlatRoundnessTop=   4
      FlatRoundnessBottom=   4
      FlatRoundnessTabs=   4
      TabMousePointerHand=   -1  'True
      HighlightMode   =   196
      HighlightModeTabSel=   64
      FlatBorderMode  =   1
      FlatBarHeight   =   5
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(20)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   20
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   21
      Left            =   240
      TabIndex        =   21
      Top             =   6600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   770
      Themed          =   0   'False
      BackColorTabs   =   12766860
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   7699508
      FlatBarColorInactive=   12766860
      FlatTabsSeparationLineColor=   -2147483633
      FlatBodySeparationLineColor=   7699508
      FlatBorderColor =   7699508
      HighlightColor  =   12766860
      HighlightColorTabSel=   -2147483633
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7699508
      FlatRoundnessTop=   4
      FlatRoundnessBottom=   4
      FlatRoundnessTabs=   4
      TabMousePointerHand=   -1  'True
      HighlightMode   =   196
      HighlightModeTabSel=   64
      FlatBorderMode  =   1
      FlatBarHeight   =   5
      FlatBarGripHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(21)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   21
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   22
      Left            =   3350
      TabIndex        =   22
      Top             =   6600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   770
      Themed          =   0   'False
      BackColorTabs   =   12766860
      BackColorTabSel =   16777215
      FlatBarColorHighlight=   7699508
      FlatBarColorInactive=   12766860
      FlatTabsSeparationLineColor=   -2147483633
      FlatBodySeparationLineColor=   7699508
      FlatBorderColor =   7699508
      HighlightColor  =   12766860
      HighlightColorTabSel=   -2147483633
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7699508
      FlatRoundnessTop=   4
      FlatRoundnessBottom=   4
      FlatRoundnessTabs=   4
      TabMousePointerHand=   -1  'True
      HighlightMode   =   196
      HighlightModeTabSel=   64
      FlatBorderMode  =   1
      FlatBarHeight   =   5
      FlatBarGripHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(22)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   22
         Left            =   120
         TabIndex        =   54
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   23
      Left            =   6460
      TabIndex        =   23
      Top             =   6600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   16777215
      ForeColorTabSel =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   855
      Themed          =   0   'False
      BackColorTabs   =   10646340
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   4228799
      FlatBarColorInactive=   10646340
      FlatTabsSeparationLineColor=   7434609
      FlatBodySeparationLineColor=   7434609
      FlatBorderColor =   -2147483633
      HighlightColor  =   10646340
      HighlightColorTabSel=   -2147483633
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   10646340
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   0
      IconColorMouseHoverTabSel=   0
      HighlightMode   =   196
      HighlightModeTabSel=   196
      FlatBarPosition =   1
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(23)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   23
         Left            =   120
         TabIndex        =   55
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   24
      Left            =   9600
      TabIndex        =   24
      Top             =   6600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   16777215
      ForeColorTabSel =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   1028
      Themed          =   0   'False
      BackColorTabs   =   10316098
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   1729514
      FlatBarColorInactive=   10316098
      FlatTabsSeparationLineColor=   7171437
      FlatBodySeparationLineColor=   7171437
      FlatBorderColor =   -2147483633
      HighlightColor  =   10316098
      HighlightColorTabSel=   -2147483633
      TabWidthStyle   =   1
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   10316098
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   0
      IconColorMouseHoverTabSel=   0
      HighlightMode   =   196
      HighlightModeTabSel=   196
      FlatBarPosition =   1
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(24)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   24
         Left            =   120
         TabIndex        =   56
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   25
      Left            =   12708
      TabIndex        =   25
      Top             =   6600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   16777215
      ForeColorTabSel =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   1028
      Themed          =   0   'False
      BackColorTabs   =   8804407
      BackColorTabSel =   -2147483633
      FlatBarColorHighlight=   1729514
      FlatBarColorInactive=   8804407
      FlatTabsSeparationLineColor=   6184542
      FlatBodySeparationLineColor=   6184542
      FlatBorderColor =   -2147483633
      HighlightColor  =   8804407
      HighlightColorTabSel=   -2147483633
      TabWidthStyle   =   1
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   8804407
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   0
      IconColorMouseHoverTabSel=   0
      HighlightMode   =   196
      HighlightModeTabSel=   196
      FlatBarPosition =   1
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(25)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   25
         Left            =   120
         TabIndex        =   57
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   26
      Left            =   15816
      TabIndex        =   26
      Top             =   6600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   16777215
      ForeColorTabSel =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   1028
      Themed          =   0   'False
      BackColorTabs   =   7742876
      BackColorTabSel =   16775410
      FlatBarColorHighlight=   5751007
      FlatBarColorInactive=   7742876
      FlatTabsSeparationLineColor=   7829367
      FlatBodySeparationLineColor=   7829367
      FlatBorderColor =   -2147483633
      HighlightColor  =   7742876
      HighlightColorTabSel=   -2147483633
      TabWidthStyle   =   1
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   7742876
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   0
      IconColorMouseHoverTabSel=   0
      HighlightMode   =   196
      HighlightModeTabSel=   196
      FlatBarPosition =   1
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(26)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFF8F2&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   26
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   27
      Left            =   18936
      TabIndex        =   27
      Top             =   6600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColor       =   16777215
      ForeColorTabSel =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   1028
      Themed          =   0   'False
      BackColorTabs   =   6955405
      BackColorTabSel =   16775410
      FlatBarColorHighlight=   1283056
      FlatBarColorInactive=   6955405
      FlatTabsSeparationLineColor=   7039851
      FlatBodySeparationLineColor=   7039851
      FlatBorderColor =   -2147483633
      HighlightColor  =   6955405
      HighlightColorTabSel=   -2147483633
      TabWidthStyle   =   1
      TabSeparation   =   8
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   6955405
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   0
      IconColorMouseHoverTabSel=   0
      HighlightMode   =   196
      HighlightModeTabSel=   196
      FlatBarPosition =   1
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(27)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFF8F2&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   27
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   28
      Left            =   240
      TabIndex        =   28
      Top             =   8760
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColorTabSel =   16777215
      ForeColorHighlighted=   16731706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   643
      Themed          =   0   'False
      BackColorTabs   =   16250871
      FlatBarColorHighlight=   16250871
      FlatBarColorInactive=   16250871
      FlatTabsSeparationLineColor=   16250871
      FlatBodySeparationLineColor=   14869218
      FlatBorderColor =   16250871
      HighlightColor  =   16477710
      HighlightColorTabSel=   16477710
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   16731706
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   16777215
      IconColorMouseHover=   16731706
      IconColorMouseHoverTabSel=   16777215
      IconColorTabHighlighted=   16731706
      HighlightMode   =   32
      HighlightModeTabSel=   20
      FlatBarHeight   =   4
      FlatBarPosition =   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(28)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   28
         Left            =   120
         TabIndex        =   60
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   29
      Left            =   3350
      TabIndex        =   29
      Top             =   8760
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      ForeColorTabSel =   16777215
      ForeColorHighlighted=   16731706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   643
      Themed          =   0   'False
      BackColorTabs   =   16250871
      FlatBarColorHighlight=   16250871
      FlatBarColorInactive=   16250871
      FlatTabsSeparationLineColor=   16250871
      FlatBodySeparationLineColor=   14869218
      FlatBorderColor =   16731706
      HighlightColor  =   16477710
      HighlightColorTabSel=   16477710
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   16731706
      FlatRoundnessTabs=   8
      TabMousePointerHand=   -1  'True
      IconColorTabSel =   16777215
      IconColorMouseHover=   16731706
      IconColorMouseHoverTabSel=   16777215
      IconColorTabHighlighted=   16731706
      HighlightMode   =   16
      HighlightModeTabSel=   20
      FlatBorderMode  =   1
      FlatBarHeight   =   4
      FlatBarPosition =   1
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(29)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   29
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   30
      Left            =   6460
      TabIndex        =   30
      Top             =   8760
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BackColor       =   16250871
      ForeColor       =   0
      FlatTabBorderColorHighlight=   8603431
      FlatTabBorderColorTabSel=   15492185
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   643
      Themed          =   0   'False
      BackColorTabs   =   16250871
      FlatTabsSeparationLineColor=   16777215
      FlatBodySeparationLineColor=   16777215
      FlatBorderColor =   16777215
      HighlightColor  =   15461355
      AutoTabHeight   =   -1  'True
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      FlatRoundnessTabs=   4
      TabMousePointerHand=   -1  'True
      HighlightMode   =   12
      HighlightModeTabSel=   512
      FlatBorderMode  =   1
      FlatBarHeight   =   0
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(30)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   30
         Left            =   120
         TabIndex        =   62
         Top             =   1080
         Width           =   2772
      End
   End
   Begin NewTabCtl.NewTab NewTab1 
      Height          =   2052
      Index           =   31
      Left            =   9564
      TabIndex        =   31
      Top             =   8760
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3620
      BackColor       =   3476744
      ForeColor       =   16777215
      FlatTabBorderColorHighlight=   14195837
      FlatTabBorderColorTabSel=   14195837
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   3
      TabHeight       =   643
      Themed          =   0   'False
      BackColorTabs   =   3476744
      BackColorTabSel =   592137
      FlatTabsSeparationLineColor=   3476744
      FlatBorderColor =   3476744
      HighlightColor  =   4210752
      AutoTabHeight   =   -1  'True
      FlatBarColorTabSel=   12335619
      FlatRoundnessTop=   0
      FlatRoundnessBottom=   0
      FlatRoundnessTabs=   4
      TabMousePointerHand=   -1  'True
      HighlightMode   =   12
      HighlightModeTabSel=   512
      FlatBorderMode  =   1
      FlatBarHeight   =   0
      FlatBodySeparationLineHeight=   0
      TabCaption(0)   =   "Tab 0"
      Tab(0).ControlCount=   1
      Tab(0).Control(0)=   "Label1(31)"
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlCount=   0
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00090909&
         Caption         =   "Theme name"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   31
         Left            =   120
         TabIndex        =   63
         Top             =   1080
         Width           =   2772
      End
   End
End
Attribute VB_Name = "frmAllThemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim c As Long
    
    For c = NewTab1.LBound To NewTab1.UBound
        Label1(c).Caption = NewTab1(c).Theme
    Next
End Sub
