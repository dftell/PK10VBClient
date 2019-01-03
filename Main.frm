VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "股票择时策略交易终端"
   ClientHeight    =   10245
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   17595
   ForeColor       =   &H80000011&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   17595
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   79
      Top             =   9900
      Width           =   17595
      _ExtentX        =   31036
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13970
            MinWidth        =   13970
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   11642
            MinWidth        =   11642
            TextSave        =   "9:39"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer_NewestData 
      Interval        =   30000
      Left            =   11160
      Top             =   150
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9675
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   17465
      _ExtentX        =   30798
      _ExtentY        =   17066
      _Version        =   393216
      Style           =   1
      Tabs            =   14
      TabsPerRow      =   14
      TabHeight       =   520
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      TabCaption(0)   =   "概况"
      TabPicture(0)   =   "Main.frx":10CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSChart1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frame_indur"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Timer_HistoryData"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Timer_Exchange"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frame_AllCols"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frame_colChance"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frame_CurrInfo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Timer_Research"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Timer_Form"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Timer_Wx"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "WebBrowser1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "交易接口"
      TabPicture(1)   =   "Main.frx":10E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Timer_WXMsg"
      Tab(1).Control(1)=   "Frame_Msg"
      Tab(1).Control(2)=   "Frame_Exchange"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "研究"
      TabPicture(2)   =   "Main.frx":1102
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dg_Research"
      Tab(2).Control(1)=   "CommonDialog1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "最新数据"
      TabPicture(3)   =   "Main.frx":111E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "dg_NewestData"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "历史数据"
      TabPicture(4)   =   "Main.frx":113A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dg_HistoryData"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "交易记录"
      TabPicture(5)   =   "Main.frx":1156
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dg_ExchangeList"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "交易汇总"
      TabPicture(6)   =   "Main.frx":1172
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "dg_ExchangeSummary"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "配置"
      TabPicture(7)   =   "Main.frx":118E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame_Config_Asset"
      Tab(7).Control(1)=   "btn_saveconfig"
      Tab(7).Control(2)=   "Frame_Config_Research"
      Tab(7).Control(3)=   "Frame_Config_System"
      Tab(7).Control(4)=   "Frame_ExchangeConfig"
      Tab(7).ControlCount=   5
      TabCaption(8)   =   "网络"
      TabPicture(8)   =   "Main.frx":11AA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "WebBrowser_set"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "图表"
      TabPicture(9)   =   "Main.frx":11C6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "3"
      TabPicture(10)  =   "Main.frx":11E2
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "MSFlexGrid6"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "4"
      TabPicture(11)  =   "Main.frx":11FE
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "MSFlexGrid7"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "5"
      TabPicture(12)  =   "Main.frx":121A
      Tab(12).ControlEnabled=   0   'False
      Tab(12).ControlCount=   0
      TabCaption(13)  =   "6"
      TabPicture(13)  =   "Main.frx":1236
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "dg_query"
      Tab(13).Control(1)=   "txt_execSql"
      Tab(13).Control(2)=   "btn_exec"
      Tab(13).Control(3)=   "btn_query"
      Tab(13).ControlCount=   4
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   7305
         Left            =   240
         TabIndex        =   177
         Top             =   2190
         Width           =   17085
         ExtentX         =   30136
         ExtentY         =   12885
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.CommandButton btn_query 
         Caption         =   "查询"
         Height          =   375
         Left            =   -59400
         TabIndex        =   131
         Top             =   450
         Width           =   765
      End
      Begin VB.CommandButton btn_exec 
         Caption         =   "执行"
         Height          =   375
         Left            =   -58620
         TabIndex        =   129
         Top             =   450
         Width           =   765
      End
      Begin VB.TextBox txt_execSql 
         Height          =   2235
         Left            =   -74880
         TabIndex        =   128
         Top             =   1050
         Width           =   17265
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser_set 
         Height          =   7035
         Left            =   -74880
         TabIndex        =   123
         Top             =   420
         Width           =   13215
         ExtentX         =   23310
         ExtentY         =   12409
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -71850
         Top             =   3180
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer_WXMsg 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   -64770
         Top             =   90
      End
      Begin VB.Frame Frame_Config_Asset 
         Caption         =   "资金设置"
         Height          =   3135
         Left            =   -61260
         TabIndex        =   85
         Top             =   2250
         Width           =   3675
         Begin VB.TextBox txt_Config_Asset_TotalCash 
            Height          =   315
            Left            =   1410
            TabIndex        =   112
            Top             =   1290
            Width           =   1755
         End
         Begin VB.TextBox txt_Config_Asset_TotalMaxRate 
            Height          =   285
            Left            =   1410
            TabIndex        =   98
            Top             =   1920
            Width           =   1755
         End
         Begin VB.TextBox txt_Config_Asset_AChanceMaxRate 
            Height          =   285
            Left            =   1410
            TabIndex        =   94
            Top             =   1620
            Width           =   1755
         End
         Begin VB.TextBox txt_Config_Asset_Gained 
            Height          =   315
            Left            =   1410
            TabIndex        =   92
            Top             =   960
            Width           =   1755
         End
         Begin VB.TextBox txt_Config_Asset_Costed 
            Height          =   315
            Left            =   1410
            TabIndex        =   90
            Top             =   630
            Width           =   1755
         End
         Begin VB.TextBox txt_Config_Asset_InitCash 
            Height          =   315
            Left            =   1410
            TabIndex        =   88
            Top             =   300
            Width           =   1755
         End
         Begin VB.Label Label38 
            Caption         =   "资金余额"
            Height          =   255
            Left            =   570
            TabIndex        =   113
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label31 
            Caption         =   "%"
            Height          =   255
            Left            =   3240
            TabIndex        =   99
            Top             =   1950
            Width           =   225
         End
         Begin VB.Label Label30 
            Caption         =   "整体止损比例"
            Height          =   225
            Left            =   240
            TabIndex        =   97
            Top             =   1920
            Width           =   1185
         End
         Begin VB.Label Label29 
            Caption         =   "%"
            Height          =   255
            Left            =   3240
            TabIndex        =   96
            Top             =   1650
            Width           =   255
         End
         Begin VB.Label Label27 
            Caption         =   "单机会使用比例"
            Height          =   255
            Left            =   60
            TabIndex        =   93
            Top             =   1620
            Width           =   1275
         End
         Begin VB.Label Label26 
            Caption         =   "当日实现盈利"
            Height          =   195
            Left            =   210
            TabIndex        =   91
            Top             =   1050
            Width           =   1125
         End
         Begin VB.Label Label25 
            Caption         =   "当日使用资金"
            Height          =   255
            Left            =   210
            TabIndex        =   89
            Top             =   690
            Width           =   1125
         End
         Begin VB.Label Label24 
            Caption         =   "初始资金"
            Height          =   225
            Left            =   570
            TabIndex        =   87
            Top             =   360
            Width           =   765
         End
      End
      Begin MSFlexGridLib.MSFlexGrid dg_ExchangeSummary 
         Height          =   9195
         Left            =   -74970
         TabIndex        =   83
         Top             =   480
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   16219
         _Version        =   393216
         SelectionMode   =   1
      End
      Begin VB.Timer Timer_Wx 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   9225
         Top             =   30
      End
      Begin VB.Frame Frame_Msg 
         Caption         =   "消息"
         Height          =   5055
         Left            =   -74850
         TabIndex        =   75
         Top             =   2445
         Width           =   13185
         Begin VB.TextBox Txt_WxMsg 
            Height          =   4665
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   76
            Top             =   240
            Width           =   12945
         End
      End
      Begin VB.CommandButton btn_saveconfig 
         Caption         =   "保存"
         Height          =   495
         Left            =   -64710
         TabIndex        =   74
         Top             =   9120
         Width           =   1005
      End
      Begin VB.Frame Frame_Exchange 
         Caption         =   "交易"
         Height          =   1875
         Left            =   -74850
         TabIndex        =   71
         Top             =   540
         Width           =   13185
         Begin VB.ComboBox dll_SendMsgUser 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   122
            Top             =   600
            Width           =   4545
         End
         Begin VB.CommandButton btn_startExchange 
            Caption         =   "启动交易"
            Height          =   555
            Left            =   12030
            TabIndex        =   84
            Top             =   1260
            Width           =   1065
         End
         Begin VB.CommandButton btn_WxSetting 
            Caption         =   "设置接口"
            Height          =   525
            Left            =   12030
            TabIndex        =   81
            Top             =   720
            Width           =   1065
         End
         Begin VB.CommandButton btn_wxStart 
            Caption         =   "启动接口"
            Height          =   555
            Left            =   12030
            TabIndex        =   80
            Top             =   150
            Width           =   1065
         End
         Begin VB.ComboBox dll_MsgUser 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   270
            Width           =   4545
         End
         Begin VB.Label Label39 
            Caption         =   " 发送用户"
            Height          =   225
            Left            =   240
            TabIndex        =   121
            Top             =   630
            Width           =   825
         End
         Begin VB.Image Image1 
            Height          =   1545
            Left            =   6150
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label22 
            Caption         =   "通讯对象"
            Height          =   225
            Left            =   330
            TabIndex        =   77
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.Timer Timer_Form 
         Interval        =   60000
         Left            =   12480
         Top             =   30
      End
      Begin VB.Timer Timer_Research 
         Interval        =   60000
         Left            =   64800
         Top             =   0
      End
      Begin VB.Frame frame_CurrInfo 
         Caption         =   "最新信息"
         Height          =   555
         Left            =   240
         TabIndex        =   36
         Top             =   450
         Width           =   17145
         Begin VB.TextBox txt_CurrExpectCount 
            Enabled         =   0   'False
            Height          =   270
            Left            =   12420
            TabIndex        =   44
            Top             =   180
            Width           =   435
         End
         Begin VB.TextBox txt_lastOpenCode 
            Enabled         =   0   'False
            Height          =   270
            Left            =   8520
            TabIndex        =   43
            Top             =   180
            Width           =   2745
         End
         Begin VB.TextBox txt_LastTime 
            Enabled         =   0   'False
            Height          =   270
            Left            =   5430
            TabIndex        =   42
            Top             =   180
            Width           =   2085
         End
         Begin VB.TextBox txt_LastExpect 
            Height          =   270
            Left            =   3240
            TabIndex        =   40
            Top             =   180
            Width           =   1275
         End
         Begin VB.TextBox txt_currExpect 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1050
            TabIndex        =   38
            Top             =   180
            Width           =   1275
         End
         Begin VB.Label Label15 
            Caption         =   "当期号码"
            Height          =   225
            Left            =   7650
            TabIndex        =   46
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label14 
            Caption         =   "当日期数"
            Height          =   225
            Left            =   11520
            TabIndex        =   45
            Top             =   240
            Width           =   945
         End
         Begin VB.Label Label13 
            Caption         =   "最后时间"
            Height          =   225
            Left            =   4620
            TabIndex        =   41
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label12 
            Caption         =   "最后期数"
            Height          =   225
            Left            =   2430
            TabIndex        =   39
            Top             =   210
            Width           =   945
         End
         Begin VB.Label Label11 
            Caption         =   "当前期数"
            Height          =   225
            Left            =   150
            TabIndex        =   37
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame frame_colChance 
         Caption         =   "单列信息"
         Height          =   1485
         Left            =   14940
         TabIndex        =   33
         Top             =   8130
         Visible         =   0   'False
         Width           =   2445
         Begin BJSCSys.CSubclass CSubclass1 
            Left            =   3180
            Top             =   1260
            _ExtentX        =   1879
            _ExtentY        =   979
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   56
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   1
            Left            =   840
            TabIndex        =   55
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   2
            Left            =   1590
            TabIndex        =   54
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   3
            Left            =   2340
            TabIndex        =   53
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   4
            Left            =   3090
            TabIndex        =   52
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   5
            Left            =   3840
            TabIndex        =   51
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   6
            Left            =   4590
            TabIndex        =   50
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   7
            Left            =   5340
            TabIndex        =   49
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   8
            Left            =   6090
            TabIndex        =   48
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txt_SingleCol 
            Height          =   315
            Index           =   9
            Left            =   6840
            TabIndex        =   47
            Top             =   330
            Width           =   735
         End
         Begin MSFlexGridLib.MSFlexGrid dg_colChances 
            Height          =   4440
            Left            =   60
            TabIndex        =   34
            Top             =   690
            Visible         =   0   'False
            Width           =   7590
            _ExtentX        =   13388
            _ExtentY        =   7832
            _Version        =   393216
         End
      End
      Begin VB.Frame frame_AllCols 
         Caption         =   "综合信息"
         Height          =   1305
         Left            =   240
         TabIndex        =   31
         Top             =   8310
         Visible         =   0   'False
         Width           =   3675
         Begin MSFlexGridLib.MSFlexGrid dg_AllChances 
            Height          =   4875
            Left            =   60
            TabIndex        =   32
            Top             =   270
            Visible         =   0   'False
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   8599
            _Version        =   393216
            Rows            =   50
            Cols            =   8
            SelectionMode   =   2
            AllowUserResizing=   1
         End
      End
      Begin VB.Timer Timer_Exchange 
         Interval        =   30000
         Left            =   11820
         Top             =   30
      End
      Begin VB.Timer Timer_HistoryData 
         Interval        =   60000
         Left            =   10110
         Top             =   0
      End
      Begin VB.Frame Frame_Config_Research 
         Caption         =   "研究设置"
         Height          =   1875
         Left            =   -74970
         TabIndex        =   4
         Top             =   5670
         Width           =   17385
         Begin VB.TextBox txt_config_Research_RepeatCheckCnt 
            Height          =   285
            Left            =   4560
            TabIndex        =   138
            Top             =   660
            Width           =   1035
         End
         Begin VB.TextBox txt_config_Research_ValidOldestHistoryExpect 
            Height          =   300
            Left            =   1320
            TabIndex        =   68
            Top             =   1470
            Width           =   1005
         End
         Begin VB.TextBox txt_config_Research_StartCol 
            Height          =   285
            Left            =   1320
            TabIndex        =   67
            Top             =   1080
            Width           =   1005
         End
         Begin VB.TextBox txt_config_Research_SingleCarRepeatCnt 
            Height          =   285
            Left            =   4560
            TabIndex        =   66
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox txt_config_Research_NewestHistoryExpect 
            Height          =   285
            Left            =   1320
            TabIndex        =   62
            Top             =   660
            Width           =   1005
         End
         Begin VB.TextBox txt_config_Research_FromPage 
            Height          =   285
            Left            =   1320
            TabIndex        =   60
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label Label43 
            Caption         =   "重复号码跟踪起点"
            Height          =   195
            Left            =   3060
            TabIndex        =   137
            Top             =   690
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "可用开始期号"
            Height          =   315
            Left            =   210
            TabIndex        =   86
            Top             =   1500
            Width           =   1125
         End
         Begin VB.Label Label19 
            Caption         =   "检查开始列"
            Height          =   255
            Left            =   300
            TabIndex        =   65
            Top             =   1110
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "单号重复次数"
            Height          =   285
            Left            =   3060
            TabIndex        =   64
            Top             =   300
            Width           =   1125
         End
         Begin VB.Label Label17 
            Caption         =   "最新历史期"
            Height          =   315
            Left            =   330
            TabIndex        =   63
            Top             =   720
            Width           =   945
         End
         Begin VB.Label Label16 
            Caption         =   "历史开始页"
            Height          =   195
            Left            =   330
            TabIndex        =   61
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame_Config_System 
         Caption         =   "系统设置"
         Height          =   1815
         Left            =   -74970
         TabIndex        =   3
         Top             =   450
         Width           =   17385
         Begin VB.TextBox txt_config_System_MinChips 
            Height          =   270
            Left            =   7800
            TabIndex        =   178
            Top             =   870
            Width           =   900
         End
         Begin VB.TextBox txt_config_System_HedgeTimes 
            Height          =   285
            Left            =   9810
            TabIndex        =   176
            Top             =   1350
            Width           =   795
         End
         Begin VB.TextBox txt_config_System_JoinHedge 
            Height          =   285
            Left            =   9810
            TabIndex        =   173
            Top             =   750
            Width           =   795
         End
         Begin VB.TextBox txt_config_System_AllowHedge 
            Height          =   285
            Left            =   9810
            TabIndex        =   172
            Top             =   1050
            Width           =   795
         End
         Begin VB.TextBox txt_config_System_LoginInFrame 
            Height          =   270
            Left            =   14790
            TabIndex        =   170
            Text            =   "Text4"
            Top             =   450
            Width           =   795
         End
         Begin VB.TextBox txt_config_System_LoginInstFillOrEncode 
            Height          =   270
            Left            =   13110
            TabIndex        =   168
            Text            =   "Text4"
            Top             =   450
            Width           =   795
         End
         Begin VB.TextBox txt_config_System_LoginHostList 
            Height          =   270
            Left            =   13920
            TabIndex        =   166
            Text            =   "Text3"
            Top             =   150
            Width           =   3405
         End
         Begin VB.TextBox txt_config_System_LoginDefaultHost 
            Height          =   270
            Left            =   13110
            TabIndex        =   165
            Text            =   "Text1"
            Top             =   150
            Width           =   795
         End
         Begin VB.TextBox txt_config_System_LoginUrlModel 
            Height          =   270
            Left            =   9810
            TabIndex        =   163
            Text            =   "Text1"
            Top             =   150
            Width           =   2325
         End
         Begin VB.TextBox txt_config_System_ClientPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   11160
            PasswordChar    =   "*"
            TabIndex        =   161
            Text            =   "abcdef123"
            Top             =   450
            Width           =   975
         End
         Begin VB.TextBox txt_config_System_ClientUsername 
            Height          =   285
            Left            =   9810
            TabIndex        =   158
            Text            =   "user331"
            Top             =   450
            Width           =   825
         End
         Begin VB.TextBox txt_config_System_IsClient 
            Height          =   285
            Left            =   7800
            TabIndex        =   156
            Text            =   "0"
            Top             =   1470
            Width           =   885
         End
         Begin VB.TextBox txt_config_System_AllowExchange 
            Height          =   270
            Left            =   7800
            TabIndex        =   143
            Top             =   1170
            Width           =   900
         End
         Begin VB.TextBox txt_config_System_TotalCnt 
            Height          =   270
            Left            =   7800
            TabIndex        =   142
            Top             =   570
            Width           =   900
         End
         Begin VB.TextBox txt_config_System_Odds 
            Height          =   270
            Left            =   7800
            TabIndex        =   132
            Top             =   270
            Width           =   900
         End
         Begin VB.TextBox txt_config_System_BackColor_G 
            Height          =   315
            Left            =   4530
            TabIndex        =   58
            Top             =   1320
            Width           =   585
         End
         Begin VB.TextBox txt_config_System_BackColor_R 
            Height          =   315
            Left            =   3390
            TabIndex        =   57
            Top             =   1320
            Width           =   585
         End
         Begin VB.TextBox txt_config_System_BackColor_B 
            Height          =   315
            Left            =   5700
            TabIndex        =   59
            Top             =   1320
            Width           =   585
         End
         Begin VB.TextBox txt_config_System_URL 
            Height          =   270
            Left            =   1320
            TabIndex        =   6
            Top             =   300
            Width           =   5505
         End
         Begin VB.TextBox txt_config_System_InterVal 
            Height          =   270
            Left            =   1320
            TabIndex        =   11
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txt_config_System_MutliColMinTimes 
            Height          =   270
            Left            =   1320
            TabIndex        =   12
            Top             =   1110
            Width           =   1095
         End
         Begin VB.TextBox txt_config_System_StartCol 
            Height          =   270
            Left            =   1320
            TabIndex        =   13
            Top             =   570
            Width           =   1095
         End
         Begin VB.TextBox txt_config_System_SingleColMinTimes 
            Height          =   270
            Left            =   1320
            TabIndex        =   14
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label56 
            Caption         =   "最小投注数"
            Height          =   225
            Left            =   6690
            TabIndex        =   179
            Top             =   900
            Width           =   945
         End
         Begin VB.Label Label55 
            Caption         =   "对冲倍数"
            Height          =   225
            Left            =   9030
            TabIndex        =   175
            Top             =   1410
            Width           =   795
         End
         Begin VB.Label Label54 
            Caption         =   "参与对冲"
            Height          =   225
            Left            =   9030
            TabIndex        =   174
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label53 
            Caption         =   "对冲下注"
            Height          =   225
            Left            =   9030
            TabIndex        =   171
            Top             =   1110
            Width           =   765
         End
         Begin VB.Label Label52 
            Caption         =   "穿透框架"
            Height          =   195
            Left            =   13920
            TabIndex        =   169
            Top             =   480
            Width           =   945
         End
         Begin VB.Label Label51 
            Caption         =   "填充/编码"
            Height          =   195
            Left            =   12180
            TabIndex        =   167
            Top             =   450
            Width           =   825
         End
         Begin VB.Label Label50 
            Caption         =   "默认主机"
            Height          =   225
            Left            =   12270
            TabIndex        =   164
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label49 
            Caption         =   "网模"
            Height          =   195
            Left            =   9330
            TabIndex        =   162
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label48 
            Caption         =   "密码"
            Height          =   195
            Left            =   10740
            TabIndex        =   160
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label47 
            Caption         =   "账号"
            Height          =   285
            Left            =   9330
            TabIndex        =   159
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label46 
            Caption         =   "是否是客户端"
            Height          =   315
            Left            =   6540
            TabIndex        =   157
            Top             =   1500
            Width           =   1125
         End
         Begin VB.Label Label45 
            Caption         =   "是否允许交易"
            Height          =   225
            Left            =   6540
            TabIndex        =   144
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label Label44 
            Caption         =   "倍数"
            Height          =   255
            Left            =   7200
            TabIndex        =   141
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label41 
            Caption         =   "赔率"
            Height          =   225
            Left            =   7200
            TabIndex        =   133
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label28 
            Caption         =   "次"
            Height          =   285
            Left            =   2490
            TabIndex        =   95
            Top             =   870
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "单列检查次数"
            Height          =   195
            Left            =   210
            TabIndex        =   9
            Top             =   1410
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "多维检查次数"
            Height          =   195
            Left            =   210
            TabIndex        =   8
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "图表查看期数"
            Height          =   225
            Left            =   150
            TabIndex        =   7
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label Label4 
            Caption         =   "列开始位置"
            Height          =   255
            Left            =   390
            TabIndex        =   5
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   "新数据Url"
            Height          =   315
            Left            =   480
            TabIndex        =   10
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.Frame Frame_ExchangeConfig 
         Caption         =   "交易设置"
         Height          =   3135
         Left            =   -74970
         TabIndex        =   2
         Top             =   2250
         Width           =   13605
         Begin VB.TextBox txt_config_Exchange_SerTotal8 
            Height          =   270
            Left            =   12930
            TabIndex        =   155
            Top             =   2610
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_SerTotal7 
            Height          =   270
            Left            =   12930
            TabIndex        =   154
            Top             =   2310
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_SerTotal6 
            Height          =   270
            Left            =   12930
            TabIndex        =   153
            Top             =   2010
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_SerTotal5 
            Height          =   270
            Left            =   12930
            TabIndex        =   152
            Top             =   1710
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_SerTotal4 
            Height          =   270
            Left            =   12930
            TabIndex        =   151
            Top             =   1410
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_SerTotal3 
            Height          =   270
            Left            =   12930
            TabIndex        =   150
            Top             =   1110
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_SerTotal2 
            Height          =   270
            Left            =   12930
            TabIndex        =   149
            Top             =   810
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_SerTotal1 
            Height          =   270
            Left            =   12930
            TabIndex        =   148
            Top             =   510
            Width           =   585
         End
         Begin VB.TextBox txt_config_Exchange_Serial8 
            Height          =   270
            Left            =   2220
            TabIndex        =   135
            Top             =   2610
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor8 
            Height          =   270
            Left            =   1350
            TabIndex        =   134
            Top             =   2610
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_Serial7 
            Height          =   270
            Left            =   2220
            TabIndex        =   125
            Top             =   2310
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor7 
            Height          =   270
            Left            =   1350
            TabIndex        =   124
            Top             =   2310
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_Serial6 
            Height          =   270
            Left            =   2220
            TabIndex        =   119
            Top             =   2010
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_Serial5 
            Height          =   270
            Left            =   2220
            TabIndex        =   118
            Top             =   1710
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_Serial4 
            Height          =   270
            Left            =   2220
            TabIndex        =   117
            Top             =   1410
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_Serial3 
            Height          =   270
            Left            =   2220
            TabIndex        =   116
            Top             =   1110
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_Serial2 
            Height          =   270
            Left            =   2220
            TabIndex        =   115
            Top             =   810
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_Serial1 
            Height          =   270
            Left            =   2220
            TabIndex        =   114
            Top             =   510
            Width           =   10695
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor4 
            Height          =   270
            Left            =   1350
            TabIndex        =   108
            Top             =   1410
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor5 
            Height          =   270
            Left            =   1350
            TabIndex        =   107
            Top             =   1710
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor6 
            Height          =   270
            Left            =   1350
            TabIndex        =   106
            Top             =   2010
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor2 
            Height          =   270
            Left            =   1350
            TabIndex        =   103
            Top             =   810
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor3 
            Height          =   270
            Left            =   1350
            TabIndex        =   102
            Top             =   1110
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_MinTimesFor1 
            Height          =   270
            Left            =   1350
            TabIndex        =   101
            Top             =   510
            Width           =   795
         End
         Begin VB.TextBox txt_config_Exchange_URL 
            Height          =   285
            Left            =   1350
            TabIndex        =   73
            Top             =   210
            Width           =   12165
         End
         Begin VB.Label Label42 
            Caption         =   "8位"
            Height          =   255
            Left            =   930
            TabIndex        =   136
            Top             =   2610
            Width           =   315
         End
         Begin VB.Label Label40 
            Caption         =   "7位"
            Height          =   255
            Left            =   930
            TabIndex        =   126
            Top             =   2310
            Width           =   315
         End
         Begin VB.Label Label37 
            Caption         =   "4位"
            Height          =   255
            Left            =   930
            TabIndex        =   111
            Top             =   1440
            Width           =   315
         End
         Begin VB.Label Label36 
            Caption         =   "5位"
            Height          =   255
            Left            =   930
            TabIndex        =   110
            Top             =   1740
            Width           =   315
         End
         Begin VB.Label Label35 
            Caption         =   "6位"
            Height          =   255
            Left            =   930
            TabIndex        =   109
            Top             =   2040
            Width           =   315
         End
         Begin VB.Label Label34 
            Caption         =   "3位"
            Height          =   255
            Left            =   930
            TabIndex        =   105
            Top             =   1110
            Width           =   315
         End
         Begin VB.Label Label33 
            Caption         =   "2位"
            Height          =   255
            Left            =   930
            TabIndex        =   104
            Top             =   810
            Width           =   315
         End
         Begin VB.Label Label32 
            Caption         =   "1位"
            Height          =   255
            Left            =   930
            TabIndex        =   100
            Top             =   510
            Width           =   315
         End
         Begin VB.Label Label21 
            Caption         =   "WXUrl"
            Height          =   225
            Left            =   780
            TabIndex        =   72
            Top             =   240
            Width           =   435
         End
      End
      Begin VB.Frame frame_indur 
         Caption         =   "指令"
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   17145
         Begin VB.CommandButton btn_AddHudgChance 
            Caption         =   "H+"
            Height          =   420
            Left            =   16080
            TabIndex        =   147
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton btn_AddSelfDifChance 
            Caption         =   "C+"
            Height          =   420
            Left            =   15600
            TabIndex        =   146
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox txt_selfDifChance 
            Height          =   270
            Left            =   13800
            TabIndex        =   145
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton btn_modifMemery 
            Caption         =   "Memery"
            Height          =   315
            Left            =   12240
            TabIndex        =   140
            Top             =   480
            Width           =   795
         End
         Begin VB.CommandButton btn_Send 
            Caption         =   "Send"
            Height          =   315
            Left            =   12240
            TabIndex        =   139
            Top             =   180
            Width           =   795
         End
         Begin VB.CommandButton btn_Incr 
            Caption         =   "Incr"
            Height          =   285
            Left            =   13140
            TabIndex        =   120
            Top             =   480
            Width           =   555
         End
         Begin VB.TextBox txt_RepeatLastDate 
            Enabled         =   0   'False
            Height          =   270
            Left            =   10290
            TabIndex        =   70
            Top             =   570
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox txt_RepeateChances 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1050
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txt_RepeatECnt 
            Enabled         =   0   'False
            Height          =   270
            Left            =   6540
            TabIndex        =   30
            Top             =   570
            Visible         =   0   'False
            Width           =   2445
         End
         Begin VB.TextBox txt_RepeatCnt 
            Enabled         =   0   'False
            Height          =   270
            Left            =   4860
            TabIndex        =   28
            Top             =   570
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton btn_copy 
            Caption         =   "Copy"
            Height          =   315
            Left            =   13140
            TabIndex        =   19
            ToolTipText     =   "Ctrl+C"
            Top             =   180
            Width           =   555
         End
         Begin VB.TextBox txt_SendTxt 
            Height          =   300
            Left            =   1050
            TabIndex        =   18
            Top             =   210
            Width           =   8595
         End
         Begin VB.TextBox txt_DataStatus 
            Height          =   300
            Left            =   11580
            TabIndex        =   16
            Text            =   "0"
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label20 
            Caption         =   "前一实现时间"
            Height          =   225
            Left            =   9090
            TabIndex        =   69
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "曾经记录"
            Height          =   225
            Left            =   5700
            TabIndex        =   29
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "当前连续出现次数"
            Height          =   225
            Left            =   3300
            TabIndex        =   27
            Top             =   600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Label8 
            Caption         =   "重复号码"
            Height          =   225
            Left            =   150
            TabIndex        =   26
            Top             =   660
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "发送指令"
            Height          =   225
            Left            =   150
            TabIndex        =   17
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label Label6 
            Caption         =   "发送状态"
            Height          =   195
            Left            =   10500
            TabIndex        =   15
            Top             =   300
            Width           =   945
         End
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid6 
         Height          =   7065
         Left            =   -74970
         TabIndex        =   20
         Top             =   480
         Width           =   13305
         _ExtentX        =   23469
         _ExtentY        =   12462
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid7 
         Height          =   7065
         Left            =   -74970
         TabIndex        =   21
         Top             =   480
         Width           =   13305
         _ExtentX        =   23469
         _ExtentY        =   12462
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid8 
         Height          =   7065
         Left            =   -74970
         TabIndex        =   22
         Top             =   480
         Width           =   13305
         _ExtentX        =   23469
         _ExtentY        =   12462
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid dg_HistoryData 
         Height          =   9165
         Left            =   -74970
         TabIndex        =   23
         Top             =   480
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   16166
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid dg_Research 
         Height          =   9165
         Left            =   -74970
         TabIndex        =   24
         Top             =   480
         Width           =   17385
         _ExtentX        =   30665
         _ExtentY        =   16166
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid dg_NewestData 
         Height          =   9195
         Left            =   -74970
         TabIndex        =   25
         Top             =   480
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   16219
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid dg_ExchangeList 
         Height          =   9195
         Left            =   -74970
         TabIndex        =   82
         Top             =   480
         Width           =   17445
         _ExtentX        =   30771
         _ExtentY        =   16219
         _Version        =   393216
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   405
         Left            =   1770
         OleObjectBlob   =   "Main.frx":1252
         TabIndex        =   127
         Top             =   2700
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSFlexGridLib.MSFlexGrid dg_query 
         Height          =   6285
         Left            =   -74940
         TabIndex        =   130
         Top             =   3300
         Width           =   17355
         _ExtentX        =   30612
         _ExtentY        =   11086
         _Version        =   393216
      End
   End
   Begin VB.Menu Menu_System 
      Caption         =   "系统"
      Begin VB.Menu menu_running 
         Caption         =   "运行"
      End
      Begin VB.Menu menu_Stop 
         Caption         =   "停止"
      End
   End
   Begin VB.Menu operate 
      Caption         =   "刷新"
      Begin VB.Menu menu_RefreshSummary 
         Caption         =   "概要"
      End
      Begin VB.Menu menu_getNewestData 
         Caption         =   "最新数据"
      End
      Begin VB.Menu menu_getHistoryData 
         Caption         =   "历史数据"
      End
      Begin VB.Menu menu_Research 
         Caption         =   "研究结果"
      End
      Begin VB.Menu menu_refreshExchange 
         Caption         =   "交易记录"
      End
      Begin VB.Menu menu_ExchangeSummary 
         Caption         =   "交易汇总"
      End
      Begin VB.Menu menu_chart 
         Caption         =   "走势图"
      End
   End
   Begin VB.Menu menu_Operate 
      Caption         =   "操作"
      Begin VB.Menu menu_ExecExchange 
         Caption         =   "执行交易"
      End
      Begin VB.Menu menu_DeleteExchangeRec 
         Caption         =   "删除当日交易"
      End
      Begin VB.Menu menu_deleteCurrDayData 
         Caption         =   "删除当日数据"
      End
      Begin VB.Menu testExchange 
         Caption         =   "启动交易"
      End
      Begin VB.Menu bucktest 
         Caption         =   "执行回测"
      End
   End
   Begin VB.Menu menu_R 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu menu_grid_Expect 
         Caption         =   "导出"
      End
      Begin VB.Menu menu_Delete 
         Caption         =   "删除"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Gobalobj As SystemClass
Public cm As New ClassModel
Public Wxobj As WXUtils
Private Const MIM_BACKGROUND = &H2
Private Const MIM_APPLYTOSUBMENUS = &H80000000
Public ExchangeForm As Form
Dim mySort As Integer
Private Type MENUINFO
    cbSize   As Long
    fMask   As Long
    dwStyle   As Long
    cyMax   As Long
    hbrBack   As Long
    dwContextHelpID   As Long
    dwMenuData   As Long
End Type
   
Private Declare Function GetMenuInfo _
                Lib "user32" (ByVal hMenu As Long, _
                              mi As MENUINFO) As Long
Private Declare Function SetMenuInfo _
                Lib "user32" (ByVal hMenu As Long, _
                              mi As MENUINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const HistoryTimes = 4
Const NewestTimes = 2
Const ResearchTimes = 2
Dim HistoryTimeCnt As Integer
Dim NewestTimesCnt As Integer
Dim ResearchTimeCnt As Integer
Dim ExportGrid As MSFlexGrid
Dim commdb As New DBClass
Dim CurrHostName As String
Dim strurl  As String '= "https://www.kcai{host}.com"
Public doc As HTMLDocument
Dim strHost As String

Private Sub btn_AddHudgChance_Click()
    On Error Resume Next
    Dim basecnt As Integer
    basecnt = CInt(InputBox("基础数量：", "对冲组合", "500"))
    Dim ccs() As ChanceClass
    Dim ccArr() As String
    If Len(Trim(Me.txt_selfDifChance.Text)) = 0 Then Exit Sub
    ccArr = Split(Trim(Me.txt_selfDifChance.Text), " ")
    Dim i As Integer
    ReDim ccs(UBound(ccArr) + 1)
    For i = 0 To UBound(ccArr)
        Set ccs(i + 1) = New ChanceClass
        ccs(i + 1).ChanceType = 2
        ccs(i + 1).ExpectCode = Me.Gobalobj.lastExpect
        ccs(i + 1).ChanceCode = ccArr(i)
        Dim strRev As String
        ccs(i + 1).InputTimes = 1
        ccs(i + 1).BaseCost = basecnt
    Next
    Dim msg As String
    msg = Me.cm.getChips(ccs, True)
    Me.SendMsg Me.Gobalobj.lastExpect, msg
End Sub

Private Sub btn_AddSelfDifChance_Click()
    Dim ccs() As ChanceClass
    Dim ccArr() As String
    If Len(Trim(Me.txt_selfDifChance.Text)) = 0 Then Exit Sub
    ccArr = Split(Trim(Me.txt_selfDifChance.Text), " ")
    Dim i As Integer
    ReDim ccs(UBound(ccArr) + 1)
    For i = 0 To UBound(ccArr)
        Set ccs(i + 1) = New ChanceClass
        ccs(i + 1).ChanceType = 0
        ccs(i + 1).ExpectCode = Me.Gobalobj.lastExpect
        ccs(i + 1).ChanceCode = ccArr(i)
        ccs(i + 1).InputTimes = 1

    Next
    Dim msg As String
    msg = Me.cm.getChips(ccs, True)
    Me.SendMsg Me.Gobalobj.lastExpect, msg
End Sub

Private Sub btn_exec_Click()
    Dim sql As String
    sql = txt_execSql.Text
    Dim cnt As Integer
    commdb.ExcelSql sql, cnt
    MsgBox Replace("修改了{0}行！", "{0}", cnt)
End Sub

Private Sub btn_Incr_Click()
    Dim data() As ExpectData
    Gobalobj.NoHtmlGetNewData = True
    cm.currExpect = CLng(Me.txt_LastExpect.Text) + 1
    cm.RefreshNewestData data, Gobalobj.NoHtmlGetNewData
    menu_getNewestData_Click
    Timer_NewestData_Timer
End Sub

Private Sub btn_modifMemery_Click()
    On Error Resume Next
    Me.cm.FirstRepeatCnt = CInt(InputBox("输入开始的次数", "修改内存", CStr(Me.cm.FirstRepeatCnt)))
End Sub

Private Sub btn_query_Click()
    Dim sql As String
    sql = txt_execSql.Text
    Dim dt As DataTable
    Set dt = commdb.getDataBySql(sql)
   Dim dgc  As New DataGridClass
   Set dgc.grid = Me.dg_query
   dgc.FillGrid dt
End Sub

Private Sub btn_saveconfig_Click()
        Me.MousePointer = vbHourglass
        Gobalobj.SaveConfig Me
        'Form_Load
        Me.SSTab1.Tab = 0
        Me.MousePointer = 1
End Sub

Private Sub btn_send_Click()
    Me.SendMsg Me.Gobalobj.lastExpect, Me.txt_SendTxt.Text
End Sub

Private Sub btn_startExchange_Click()
    If Gobalobj.ExchangeStatus = False Then
        btn_startExchange.Caption = "停止交易"
        Gobalobj.ExchangeStatus = True
    Else
        btn_startExchange.Caption = "启动交易"
        Gobalobj.ExchangeStatus = False
    End If
End Sub



Private Sub bucktest_Click()
    Dim frm As New frm_BackTest
    Set frm.gobj = Me.Gobalobj
    frm.Show
End Sub

Private Sub dg_ExchangeList_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    Set ExportGrid = Nothing
    If Button = 2 Then
        Set ExportGrid = Me.dg_ExchangeList
        Me.dg_ExchangeList.row = Me.dg_ExchangeList.MouseRow
        PopupMenu Me.menu_R
    End If
End Sub

Private Sub dg_ExchangeSummary_DblClick()
    Dim R As Long
    With dg_ExchangeSummary
        R = .MouseRow
        If R > .FixedRows And R < .Rows Then
            Dim frm As New Form3
            Set frm.Gobalobj = Me.Gobalobj
            frm.FrmType = "Chance"
            frm.Params = .TextMatrix(R, 1)
            frm.Caption = .TextMatrix(R, 1) & "交易明细"
            frm.Show
        End If
    End With
End Sub

Private Sub dg_Research_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    Set ExportGrid = Nothing
    With dg_Research
    If Button = 1 And .MouseRow = 0 Then
        If mySort = 0 Then
            mySort = 3
        Else
            mySort = 7 - mySort
        End If
        .col = .MouseCol '选择排序列
        '.ColSel = 1
        .Sort = mySort '排序方式
    End If
    End With
    If Button = 2 Then
        Set ExportGrid = Me.dg_Research
        PopupMenu Me.menu_R
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Me.ExchangeForm = Nothing
End Sub

Private Sub menu_chart_Click()
    Exit Sub
    Dim wclass As New WaveClass
    wclass.CheckDays = Me.Gobalobj.InterVal
    'wclass.init Gobalobj.CurrExpectData
    Dim data() As ExpectData
    Dim AllData() As ExpectData
    AllData = Gobalobj.CurrExpectData
    If UBound(Gobalobj.CurrExpectData) = 0 Then Exit Sub
    Dim cnt As Integer
    cnt = UBound(AllData)
    DoEvents
    With Me.MSChart1
               
        '.Plot.axis(VtChAxisIdY).ValueScale.Auto = False
        .columnCount = 1
        .Column = 1
        If cnt - wclass.CheckDays > 0 Then
            .RowCount = cnt - wclass.CheckDays
        Else
            Me.Gobalobj.LogObj.Log "数据数小于最小数量"
            Exit Sub
        End If
       
        Dim i As Integer
        Dim maxval As Single
        Dim minval As Single
        maxval = 0
        minval = 999
        Dim val() As Single
        ReDim val(cnt - wclass.CheckDays)
        For i = 0 To cnt - wclass.CheckDays - 1
            data = wclass.getData(AllData, cnt - wclass.CheckDays - i, wclass.CheckDays)
            Dim result As Dictionary
            Set result = wclass.CalcChartData(data)
            val(i) = result(0)
            If result(0) > maxval Then maxval = result(0)
            If result(0) < minval Then minval = result(0)

        Next
        .Plot.axis(VtChAxisIdY).ValueScale.Maximum = maxval - minval '设置纵轴标注最大值
        .Plot.axis(VtChAxisIdY).ValueScale.Minimum = 0  '设置纵轴标注最大值
        For i = 0 To cnt - wclass.CheckDays - 1
            .row = i + 1
            .data = val(i) - minval
            .RowLabel = TimeSerial(Hour(data(1).OpenTime), Minute(data(1).OpenTime), Second(data(1).OpenTime))
        Next
        
    End With
End Sub

Private Sub menu_Delete_Click()
    If ExportGrid Is Nothing Then Exit Sub
    If ExportGrid Is Me.dg_ExchangeList Then
        Dim cc As New ChanceClass
        Dim cindex As String
        cindex = Me.dg_ExchangeList.TextMatrix(Me.dg_ExchangeList.row, 1)
        If MsgBox("你是否确实要删除记录" & cindex & "?", vbYesNo, "删除交易记录") = vbYes Then
            cc.DeleteChance cindex
            menu_refreshExchange_Click
        End If
        
    End If
End Sub

Private Sub menu_grid_Expect_Click()
    ExportFlexDataToExcel ExportGrid, Me.CommonDialog1
    Set ExportGrid = Nothing
End Sub

Public Function ExportFlexDataToExcel(flex As MSFlexGrid, g_CommonDialog As CommonDialog)
  On Error GoTo ErrHandler
  Dim xlApp As Object
  Dim xlBook As Object
  Dim Rows As Integer, Cols As Integer
  Dim iRow As Integer, hCol As Integer, iCol As Integer
  Dim New_Col As Boolean
  Dim New_Column As Boolean
  g_CommonDialog.CancelError = True
  On Error GoTo ErrHandler
  ' 设置标志
  g_CommonDialog.Flags = cdlOFNHideReadOnly
  ' 设置过滤器
  g_CommonDialog.Filter = "All Files (*.*)|*.*|Excel Files(*.xls)|*.xls|Batch Files (*.bat)|*.bat"
  ' 指定缺省的过滤器
  g_CommonDialog.FilterIndex = 2
  ' 显示“打开”对话框
  g_CommonDialog.ShowSave
  If flex.Rows <= 1 Then       '判断表格中是否有数据
         MsgBox "没有数据！", vbInformation, "警告"
        Exit Function
  End If
   '打开Excel ,添加工作簿
  Set xlApp = CreateObject("Excel.Application")
  Set xlBook = xlApp.Workbooks.Add
  xlApp.Visible = False
'遍历表格中的记录，传递到Excel中
  With flex
    Rows = .Rows
    Cols = .Cols
    iRow = 0
    iCol = 1
    Me.MousePointer = vbHourglass
    For hCol = 0 To Cols - 1
       For iRow = 1 To Rows
'获取表格中值，传递到Excel单元格中
          xlApp.cells(iRow, iCol).Value = .TextMatrix(iRow - 1, hCol)
       Next iRow
       iCol = iCol + 1
    Next hCol
    Me.MousePointer = 1
  End With
   '设置Excel的属性
  With xlApp
  .Rows(1).Font.Bold = True
  .cells.Select
  .Columns.AutoFit
  .cells(1, 1).Select
' .Application.Visible = True
  End With
  '获取要保存文件的文件名，选择保存路径
  xlBook.SaveAs (g_CommonDialog.FileName)
  xlApp.Application.Visible = True
  xlApp.DisplayAlerts = False
  Set xlApp = Nothing '"交还控制给Excel
  Set xlBook = Nothing
  MsgBox "数据已经导出到Excel中。", vbInformation, "成功"
 Exit Function
      
  
ErrHandler:
  
  ' 用户按了“取消”按钮
  If Err.Number <> 32755 Then
    MsgBox "数据导出失败！", vbCritical, "警告"
  End If
End Function


Private Sub Form_DblClick()
   'Me.Hide
End Sub

Private Sub Form_Load()
   
    On Error Resume Next
    If App.PrevInstance = True Then
        MsgBox "该程序已经运行，请退出！"
        End
    End If
    Set Gobalobj = New SystemClass
    Set commdb.gobj = Me.Gobalobj
    Me.Timer_NewestData.InterVal = 10000
    Me.Caption = "策略交易终端[" & Gobalobj.ClientUserName & "]"
    Gobalobj.LogObj.Log "初始化系统....."
     '禁用所有计时器
    Me.Timer_Exchange.Enabled = False
    Me.Timer_Form.Enabled = False
    Me.Timer_HistoryData.Enabled = False
    Me.Timer_NewestData.Enabled = False
    Me.Timer_Research.Enabled = False
    Me.Timer_Wx.Enabled = False
    Me.Timer_WXMsg.Enabled = False
    Gobalobj.LogObj.Log "停止计时!"
    Err.Clear
    Gobalobj.LogObj.Log "填充控件....."
    Gobalobj.FillControl Me
    If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    Set cm.gobj = Gobalobj
    Gobalobj.LogObj.Log "初始化参数....."
    cm.InitParams Gobalobj
    If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    
    Me.BackColor = Gobalobj.BackColor
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        ctrl.BackColor = Gobalobj.BackColor
    Next
    
    
    Set Wxobj = New WXUtils
'    For I = 0 To Me.SSTab1.Tabs
'       SSTab1.Visible = False
'    Next
    CSubclass1.SubClassMe SSTab1.hwnd, 0, , Gobalobj.BackColor
     If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    Dim MyMenu As MENUINFO
    MyMenu.cbSize = Len(MyMenu)
    MyMenu.fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
    MyMenu.hbrBack = CreateSolidBrush(Gobalobj.BackColor)
     If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    Gobalobj.LogObj.Log "初始化菜单....."
    SetMenuInfo GetMenu(Me.hwnd), MyMenu
    If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    If Me.Gobalobj.IsClient Then
        '禁用多个Timer
        Me.Timer_NewestData.Enabled = True
        Me.Timer_NewestData.InterVal = 60000
        If Me.Gobalobj.ExchangeSwitched = False And Gobalobj.ClientUserName <> "" Then
            testExchange_Click
        End If
        Gobalobj.LogObj.Log "启用最新数据计时器,计时器间隔秒数为" & Me.Timer_NewestData.InterVal / 1000
        Gobalobj.LogObj.Log "系统载入完成"
        Exit Sub
    End If
    Gobalobj.LogObj.Log "初始化窗体....."
    initGrid
    If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    
    Gobalobj.LogObj.Log "初始化历史....."
    cm.InitHistoryData
     If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    Gobalobj.LogObj.Log "获取历史....."
    'menu_getHistoryData_Click
    cm.fNewestData = True '初始化时必须要设为真，后面才会触发
     If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    Gobalobj.LogObj.Log "初始化最新数据....."
    menu_getNewestData_Click
    If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    Gobalobj.LogObj.Log "刷新概要....."
    menu_RefreshSummary_Click
    If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
    End If
    Gobalobj.LogObj.Log "系统成功载入....."
    Me.WebBrowser1.Silent = True
    Me.WebBrowser1.Navigate "http://user.opencai.net/passport/login.aspx"
    Me.Timer_Exchange.Enabled = True
    Me.Timer_Form.Enabled = True
    Me.Timer_HistoryData.Enabled = True
    Me.Timer_NewestData.Enabled = True
    Me.Timer_Research.Enabled = True
    'Me.Timer_Wx.Enabled = False
    'Me.Timer_WXMsg.Enabled = False
End Sub



Private Sub menu_deleteCurrDayData_Click()
    Dim ed As New ExpectData
    ed.DeleteDayData Date
End Sub

Private Sub menu_DeleteExchangeRec_Click()
    Dim cc As New ChanceClass
    cc.DeleteDayData Date
    menu_refreshExchange_Click
End Sub

Private Sub menu_ExchangeSummary_Click()
    If Me.Gobalobj.IsClient Then Exit Sub
    
    Dim cc As New ChanceClass
    Dim dt As DataTable
    Set dt = cc.getSummary()
   Dim dgc  As New DataGridClass
   Set dgc.grid = Me.dg_ExchangeSummary
   dgc.FillGrid dt
End Sub

Private Sub menu_ExecExchange_Click()
    If Me.Gobalobj.IsClient Then Exit Sub
    If cm.ExpectInstructStatusDic.Exists(cm.currExpect) Then Exit Sub '防止多次执行
    menu_Research_Click
    Me.txt_SendTxt.Text = cm.getCurrExpectChips(Me, Me.dg_AllChances, Me.dg_colChances)
    menu_refreshExchange_Click
    If Not cm.ExpectInstructStatusDic.Exists(cm.currExpect) Then
        cm.ExpectInstructStatusDic.Add cm.currExpect, True
    End If
End Sub

Private Sub menu_getHistoryData_Click()
    If Me.Gobalobj.IsClient Then Exit Sub
   On Error Resume Next
   Dim pagecnt As Long
   
   pagecnt = Gobalobj.HistoryFromPage
   Me.Gobalobj.LogObj.Log "获取网页历史数据"
   cm.GetHistoryData pagecnt
   If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
   End If
   Me.txt_config_Research_FromPage.Text = pagecnt
   Me.txt_config_Research_NewestHistoryExpect.Text = cm.CurrHistoryExpect
   Me.Gobalobj.LogObj.Log "更新系统配置重的历史数据"
   Gobalobj.SaveConfig Me
   If Err.Number <> 0 Then
        Gobalobj.LogObj.Log Err.Description
        Err.Clear
   End If
   Dim db As New DBClass
   db.gobj = Me.Gobalobj
   Dim dt As DataTable
'''   Me.Gobalobj.LogObj.Log "读取历史数据片段"
'''   Set dt = db.getDataBySql("select top 2000 * from historydata order by expect desc")
'''   Dim dgc  As New DataGridClass
'''   Set dgc.grid = dg_HistoryData
'''   Me.Gobalobj.LogObj.Log "填充数据到表格"
'''   dgc.FillGrid dt
'''   If Err.Number <> 0 Then
'''        Gobalobj.LogObj.Log Err.Description
'''        Err.Clear
'''   End If
End Sub

Private Sub menu_getNewestData_Click()
    If Me.Gobalobj.IsClient Then

        Dim ret As String
        Dim strExpect As String
        Dim Expect As Object
        Dim strLastTime As String
        Dim strInsts As String
        If Me.Gobalobj.InstsList Is Nothing Then
            Set Me.Gobalobj.InstsList = New Dictionary
        End If
        Dim jo As New JsonClass
        ret = cm.ReReadInst()
        If ret = "" Then Exit Sub
        Dim obj As Object
        Set obj = jo.GetJsonVal(ret, "")
        strExpect = obj.Expect
        strInsts = obj.Insts
        strLastTime = obj.LastTime
        Me.Gobalobj.lastExpect = strExpect
        If Me.Gobalobj.InstsList.Exists(strExpect) Then
            Exit Sub
        End If
        Dim fullobj As Object
        Set fullobj = jo.GetJsonVal(ret, "Full")
        Dim ccss() As Object
        Dim fullInsts As String
        If (fullobj.Count <> "0") Then
            
            jo.ToArray fullobj, "ChanceList", ccss
            Dim i As Integer
            Dim txt() As String
            ReDim txt(UBound(ccss))
            For i = 1 To UBound(ccss)
                Dim ccs As Object
                Set ccs = ccss(i)
                If (ccs.ChanceCode = "") Then GoTo ThisEndFor
                If ccs.UnitCost = "" Then GoTo ThisEndFor
                If ccs.ChanceType <> 2 Then
                    If ccs.ChanceType = 1 Then
                        txt(i) = Replace(ccs.ChanceCode, "+", "/" & CLng(ccs.UnitCost) * Me.Gobalobj.SerTotal(1) & "+") & "/" & CLng(ccs.UnitCost) * Me.Gobalobj.SerTotal(1)
                    Else
                        txt(i) = Replace(ccs.ChanceCode, "+", "/" & CLng(ccs.UnitCost) * Me.Gobalobj.SerTotal(CInt(ccs.ChipCount)) & "+") & "/" & CLng(ccs.UnitCost) * Me.Gobalobj.SerTotal(CInt(ccs.ChipCount))
                    End If
                Else '如果是对冲请求
                    If Me.Gobalobj.JoinHedge = False Then '如果不参与对冲,该指令直接去除 2018/7/19
                        txt(i) = ""
                    Else
                        If Me.Gobalobj.AllowHedge Then  '如果需要下注，增加按指定对冲倍数下注,原始数量要乘以指定倍数 2018/7/19
                            'txt(i) = ccs.ChanceCode & "/" & CStr(CLng(ccs.UnitCost) + ccs.BaseCost ) & "+" & cm.getRevChance(ccs.ChanceCode) & "/" & ccs.BaseCost
                            txt(i) = ccs.ChanceCode & "/" & CStr(CLng(ccs.UnitCost * Me.Gobalobj.HedgeTimes) + ccs.BaseCost * Me.Gobalobj.HedgeTimes) & "+" & cm.getRevChance(ccs.ChanceCode) & "/" & ccs.BaseCost * Me.Gobalobj.HedgeTimes
                        Else
                            'txt(i) = Trim(ccs.ChanceCode) & "/" & ccs.UnitCost
                            txt(i) = Trim(ccs.ChanceCode) & "/" & ccs.UnitCost * Me.Gobalobj.HedgeTimes
                        End If
                    End If
                End If
ThisEndFor:
            Next
        End If
        Me.Gobalobj.InstsList.Add strExpect, strInsts
        Me.txt_currExpect.Text = strExpect
        Me.txt_LastExpect.Text = CLng(strExpect) - 1
        Me.txt_LastTime.Text = strLastTime
        Me.txt_SendTxt.Text = Replace(Replace(Trim(Join(txt, " ")), "  ", " "), "  ", " ")
        Me.Gobalobj.LogObj.Log Me.txt_SendTxt.Text
        If Me.Gobalobj.InstsList.Count > 1 Then Me.SendMsg strExpect, Me.txt_SendTxt.Text
        Exit Sub
    End If
  On Error Resume Next
  Me.StatusBar1.Panels(1).Text = "刷新最新数据..."
  Dim data() As ExpectData
  Dim bGetNewest As Boolean
  bGetNewest = cm.RefreshNewestData(data, Gobalobj.NoHtmlGetNewData)
  If bGetNewest = False And Hour(Now()) > 9 Then
    Me.StatusBar1.Panels(1).Text = "当前数据为最新."
    Me.txt_DataStatus.Text = "Waiting"
    Exit Sub
  End If
  
   Dim ed As New ExpectData
   Set ed.db.gobj = Me.Gobalobj
   Dim dt As DataTable
   Set dt = ed.getDayData(DateAdd("D", -1, Date), cm.currExpect)
   If dt.RowCount > 0 Then
        Dim dir As Dictionary
        Set dir = dt.Rows(1)
        Me.txt_LastExpect.Text = dir("Expect")
        Me.txt_lastOpenCode.Text = dir("OpenCode")
        Me.txt_LastTime.Text = dir("OpenTime")
        Me.txt_currExpect.Text = dir("Expect") + 1
        Me.txt_CurrExpectCount.Text = dt.RowCount
   End If
   Dim dgc  As New DataGridClass
   Set dgc.grid = dg_NewestData
   dgc.FillGrid dt
   Me.StatusBar1.Panels(1).Text = "获取最新数据！"
   cm.fNewestData = True  '设置当前系统的最新数据标志为真，在刷新界面后重置为否
End Sub



Private Sub menu_refreshExchange_Click()
    If Me.Gobalobj.IsClient Then Exit Sub
    'Timer_Form_Timer
    Dim cc As New ChanceClass
    Set cc.db.gobj = Me.Gobalobj
    Dim dt As DataTable
    Set dt = cc.getDayData(DateAdd("D", -1, Date))
   Dim dgc  As New DataGridClass
   Set dgc.grid = Me.dg_ExchangeList
   dgc.FillGrid dt
End Sub

Private Sub menu_RefreshSummary_Click()
    If Me.Gobalobj.IsClient Then Exit Sub
    DoEvents
    initGrid
    cm.Process dg_AllChances, Me.dg_colChances, "txt_SingleCol"
End Sub


Private Sub menu_Research_Click()
    If Me.Gobalobj.IsClient Then Exit Sub
    Me.StatusBar1.Panels(1).Text = "重新计算研究结果..."
    Dim chs() As ChanceClass
    Me.MousePointer = vbHourglass
    Dim Chances As String
    Dim holdingCnt As Integer
    Dim holdingECnt As Integer
    Dim lastDate As String
    cm.Research chs, dg_Research, Chances, holdingCnt, holdingECnt, lastDate
    If Len(Trim(chance)) > 0 Then
        Me.txt_RepeateChances.BackColor = vbWhite
    Else
        Me.txt_RepeateChances.BackColor = Gobalobj.BackColor
    End If
    'Me.dg_Research.col = 1
    'Me.dg_Research.Sort = 4
    Me.txt_RepeateChances.Text = Chances
    If Me.dg_Research.Rows > 11 Then
        Me.txt_RepeatCnt.Text = Me.dg_Research.TextMatrix(1, 4) & "/" & Me.dg_Research.TextMatrix(1, 5)
        Dim arr(10) As String
        For i = 1 To 10
            arr(i - 1) = Me.dg_Research.TextMatrix(i + 1, 4)
        Next
        Me.txt_RepeatECnt.Text = Join(arr, ",")
        Me.txt_RepeatLastDate.Text = Me.dg_Research.TextMatrix(1, 2)
    End If
    Me.txt_config_Research_ValidOldestHistoryExpect.Text = cm.ValidOldestHistoryExpect
    Me.MousePointer = 1
    Me.StatusBar1.Panels(1).Text = "研究结果为最新"
End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub menu_running_Click()
    Me.Gobalobj.AllowExchange = True
End Sub

Private Sub menu_Stop_Click()
    Me.Gobalobj.AllowExchange = False
End Sub

Private Sub SSTab1_DblClick()
    'Me.Hide
End Sub

Public Sub testExchange_Click()
    OpenTheLoginFrm
End Sub

Public Function OpenTheLoginFrm()
'''''''''    Dim frm  As Form
'''''''''    If Me.ExchangeForm Is Nothing Then
'''''''''
'''''''''        If Me.Gobalobj.IsClient Then
'''''''''            Set frm = New frm_WbSender
'''''''''        Else
'''''''''            Set frm = New frm_Test
'''''''''        End If
'''''''''        Set frm.gobj = Me.Gobalobj
'''''''''        Set frm.parentForm = Me
'''''''''        Set Me.ExchangeForm = frm
'''''''''    Else
'''''''''        Set frm = Me.ExchangeForm
'''''''''    End If
'''''''''    Set OpenTheLoginFrm = frm
'''''''''    frm.Show
Dim jo As New JsonClass
    If Me.Gobalobj.IsClient Then
            Dim strLogRet As String
            strLogRet = cm.Login()
            Dim obj As Object
            Set obj = jo.GetJsonVal(strLogRet, "")
            If (obj Is Nothing) Then
                MsgBox strLogRet
                Exit Function
            End If
            strurl = Me.Gobalobj.LoginUrlModel
            strHost = Me.Gobalobj.LoginDefaultHost
            'Me.WebBrowser1.Silent = True
             Me.WebBrowser1.Navigate Replace(strurl, "{host}", strHost)
            'Me.Caption = gobj.ClientUserName
    Else
    End If
    
    
End Function

Private Sub Timer_Exchange_Timer()
    menu_refreshExchange_Click
End Sub

Private Sub Timer_Form_Timer()
    'Me.StatusBar1.Panels(1).Text = Now()
    If Me.Gobalobj.IsClient Then Exit Sub
    If cm.fNewestData = False Then
        Me.CloseEndChances
    End If
    Gobalobj.SaveConfig Me
End Sub

Private Sub Timer_HistoryData_Timer()
    If Me.Gobalobj.IsClient Then Exit Sub
    HistoryTimeCnt = HistoryTimeCnt + 1
    If HistoryTimeCnt >= HistoryTimes And Hour(Now()) < 9 Then
        DoEvents
        HistoryTimeCnt = 0
        menu_getHistoryData_Click
    End If
End Sub

Private Sub Timer_NewestData_Timer()
    Me.Gobalobj.LogObj.Log "开始接收最新数据！"
    If Hour(Now()) < 9 And Hour(Now()) >= 1 Then Exit Sub
    DoEvents
        If (Minute(Now()) Mod 5) <= 2 Then
            Timer_NewestData.InterVal = 65535
            Exit Sub
        End If
    'menu_chart_Click
    menu_getNewestData_Click
    If Me.Gobalobj.IsClient Then
        If (Minute(Now()) Mod 5) = 3 Then
            Timer_NewestData.InterVal = 20000
        End If
        Exit Sub
    End If
    If cm.fNewestData Then
        Timer_NewestData.InterVal = 300000
        menu_Research_Click
        menu_RefreshSummary_Click
        menu_ExecExchange_Click
        cm.fNewestData = False
    Else
        If (Minute(Now()) Mod 5) = 3 Then
            Timer_NewestData.InterVal = 20000
        End If
    End If
End Sub

Function SendMsg(Expect As String, msg As String) As Boolean
    Dim ips As String
    ips = "773,331,333"
    Dim ipsArr() As String
    ipsArr = Split(ips, ",")
    Dim i As Integer
    i = 0
    Dim loadcnt As Integer
    If Me.Gobalobj.AllowExchange = False Then Exit Function
    Dim c2i As New CCS2InstrClass
    '/c2i.cJsOdds = Me.Gobalobj.Odds
    'strJson = c2i.InstrToJsonString(msg) 'Trim(Me.parentForm.cm.ToSerial(msg))
    'MsgBox strJson
    Dim frm As Form
    If Gobalobj.ExchangeSwitched = False Then
        MsgBox "请先启动交易界面！"
        Exit Function
    End If
'''    If IsNull(Me.ExchangeForm.lbl_ExecStatus) Then
'''    Else
'''        Me.ExchangeForm.lbl_ExecStatus.Caption = IIf(Me.Gobalobj.AllowExchange, "允许", "禁止")
'''    End If
    
    If Len(Trim(msg)) = 0 Then
        SendMsg = True
        Me.txt_DataStatus.Text = True
        Exit Function
    End If
    Me.txt_DataStatus.Text = False
    Dim suc As Boolean
    Dim retamt As Currency
    
    'suc = Me.ExchangeForm.SendMsg(Expect, msg, 1, retamt)
    suc = Me.RealSendMsg(Expect, msg, 1, retamt)
    If Me.Gobalobj.IsClient Then
        SendMsg = suc
        Me.StatusBar1.Panels(2).Text = retamt
        Me.txt_DataStatus.Text = suc
        Exit Function
    End If
    If suc = False Then
        While suc = False '持续发送
            DoEvents
            i = i + 1
            
            'suc = Me.ExchangeForm.SendMsg(Expect, msg, 1, retamt)
            suc = Me.RealSendMsg(Expect, msg, 1, retamt)
            Sleep 500
            If i > 5 Then
                Me.txt_DataStatus.Text = suc
                Exit Function
            End If
        Wend
    End If
    SendMsg = suc
    Me.StatusBar1.Panels(2).Text = retamt
    Me.txt_DataStatus.Text = suc
End Function

Sub initGrid()
    On Error Resume Next
    dg_AllChances.AllowUserResizing = 1
    dg_AllChances.Clear
    dg_AllChances.Rows = 0
    dg_colChances.Clear
    dg_colChances.Rows = 0
    dg_AllChances.Rows = 100
    dg_AllChances.Cols = 8
    dg_AllChances.TextMatrix(0, 1) = "2组"
    dg_AllChances.TextMatrix(0, 2) = "3组"
    dg_AllChances.TextMatrix(0, 3) = "4组"
    dg_AllChances.TextMatrix(0, 4) = "5组"
    dg_AllChances.TextMatrix(0, 5) = "6组"
    dg_AllChances.TextMatrix(0, 6) = "7组"
    dg_AllChances.TextMatrix(0, 7) = "8组"
    dg_AllChances.ColWidth(1) = 200
    dg_AllChances.ColWidth(2) = 1500
    dg_AllChances.ColWidth(3) = 1500
    dg_AllChances.ColWidth(4) = 1000
    dg_AllChances.ColWidth(5) = 1200
    dg_AllChances.ColWidth(6) = 1200
    dg_AllChances.ColWidth(7) = 1200
    dg_AllChances.ColAlignment(1) = flexAlignLeftCenter
    dg_AllChances.ColAlignment(2) = flexAlignLeftCenter
    dg_AllChances.ColAlignment(3) = flexAlignLeftCenter
    dg_AllChances.ColAlignment(4) = flexAlignLeftCenter
    dg_AllChances.ColAlignment(5) = flexAlignLeftCenter
    dg_AllChances.ColAlignment(6) = flexAlignLeftCenter
    dg_AllChances.ColAlignment(7) = flexAlignLeftCenter
    'dg_AllChances.CellBackColor = dg_AllChances.Container.Gobalobj.BackColor
    
    dg_colChances.Rows = 100
    dg_colChances.Cols = 10 + 1
    Dim i As Integer
    For i = 1 To dg_AllChances.Cols - 1
        dg_AllChances.col = i
        If i = 2 Then
            dg_AllChances.row = Gobalobj.MinTimeForChance(3) - Gobalobj.MutliColMinTimes + 1
            dg_AllChances.CellBackColor = vbWhite
            dg_AllChances.row = Gobalobj.MinTimeForChance(2) - Gobalobj.MutliColMinTimes + 1
            dg_AllChances.CellBackColor = vbWhite
            dg_AllChances.row = 37 - Gobalobj.MutliColMinTimes
            dg_AllChances.CellBackColor = vbWhite
        End If
        If i > 2 Then
            dg_AllChances.row = Gobalobj.MinTimeForChance(i + 1) - Gobalobj.MutliColMinTimes + 1
            dg_AllChances.CellBackColor = vbWhite
        End If
    Next
    With dg_colChances
    For i = 1 To .Cols - 1
        .col = i
        .row = Gobalobj.MinTimeForChance(3) - Gobalobj.MutliColMinTimes + 1
        .CellBackColor = vbWhite
        .row = Gobalobj.MinTimeForChance(4) - Gobalobj.MutliColMinTimes + 1
        .CellBackColor = vbWhite
        .row = Gobalobj.MinTimeForChance(5) - Gobalobj.MutliColMinTimes + 1
        .CellBackColor = vbWhite
        .row = Gobalobj.MinTimeForChance(6) - Gobalobj.MutliColMinTimes + 1
        .CellBackColor = vbWhite
        .row = Gobalobj.MinTimeForChance(7) - Gobalobj.MutliColMinTimes + 1
        .CellBackColor = vbWhite
        .row = Gobalobj.MinTimeForChance(2) - Gobalobj.MutliColMinTimes + 1
        .CellBackColor = vbWhite
        .row = 37 - Gobalobj.MutliColMinTimes
        .CellBackColor = vbWhite
    Next
    End With
    For i = 1 To 10
        dg_colChances.TextMatrix(0, i) = Replace("第X道", "X", i)
    Next
    For i = 1 To 80
        dg_AllChances.TextMatrix(i, 0) = Replace("X级", "X", i)
        dg_colChances.TextMatrix(i, 0) = Replace("X级", "X", i)
    Next
    With Me.dg_Research
        .Cols = 10
        .Rows = 3000
        .TextMatrix(0, 1) = "最后期次"
        .TextMatrix(0, 2) = "最后时间"
        .TextMatrix(0, 3) = "最后组合"
        .TextMatrix(0, 4) = "持续次数"
        .TextMatrix(0, 5) = "持续期数"
        .TextMatrix(0, 6) = "名词"
        .TextMatrix(0, 7) = "车次"
    End With
End Sub



Private Sub Timer_Research_Timer()
    If Me.Gobalobj.IsClient Then Exit Sub
    ResearchTimeCnt = ResearchTimeCnt + 1
    If ResearchTimeCnt >= ResearchTimes Then
        ResearchTimeCnt = 0
        menu_Research_Click
    End If
End Sub

Sub CloseEndChances()
    If Me.Gobalobj.IsClient Then Exit Sub
        Dim ccs() As ChanceClass
        cm.InitChanceList
        cm.SearchInputItems Me, Me.dg_AllChances, Me.dg_colChances, cm.chanceDic, ccs
        cm.CloseTheEndChances ccs, False  '关闭实现的机会
        
End Sub

Private Sub Timer_Wx_Timer()
    'refresh q
    Wxobj.refreshCode Me.Image1
End Sub

Private Sub btn_wxStart_Click()
    Wxobj.refreshCode Me.Image1 '刷新二维码
    Timer_Wx.Enabled = True
    Gobalobj.WXLogined = False
    Wxobj.ScanedTheQCore = False
    Timer_WXMsg.Enabled = True
End Sub

Private Sub Timer_WXMsg_Timer()
    'check
    If Gobalobj.WXLogined Then
        
    Else
        'Sheet10.Cells(4, 10) = Now()
        'UserForm1.Caption = Now()
        Me.StatusBar1.Panels(2).Text = "等待登陆..."
        Dim strCheckFlg As String
        'strCheckFlg = Sheet10.Cells(2, 1).Text
        If Wxobj.ScanedTheQCore Then Exit Sub
        Dim strurl As String
        If Wxobj.wait_login(strurl) Then
            Wxobj.ScanedTheQCore = True
            Timer_Wx.Enabled = False
            'Sheet10.Cells(2, 1) = 1
            Me.StatusBar1.Panels(2).Text = "验证成功！"
            Wxobj.init strurl, Me.StatusBar1
            Wxobj.BindToDDL Me.dll_MsgUser
            Wxobj.BindToDDL Me.dll_SendMsgUser
            Me.StatusBar1.Panels(2).Text = "控件初始化成功！"
            Gobalobj.WXLogined = True
            Exit Sub
         End If
    End If
    
End Sub





Private Sub txt_selfDifChance_DblClick()
    Me.txt_selfDifChance.Text = Me.cm.getRevChance(Me.txt_selfDifChance.Text)
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pdisp As Object, url As Variant)
    On Error Resume Next
    
    If WebBrowser1.readyState < READYSTATE_COMPLETE Then
        Exit Sub
    End If
    cm.gobj = gobj
    Dim obj As IUnknown
    Set obj = IUnknown
    If (pdisp Is Me.WebBrowser1.object) Then
    Else
        'MsgBox pDisp
    End If

    WebBrowser1.RegisterAsBrowser = True
    If WebBrowser1.LocationURL Then
        Set doc = WebBrowser1.Document
    End If
    'End If
    If doc Is Nothing Then Exit Sub
    Dim txtUser As Object
''    Set txtUser = doc.getElementById("loginName")
''    If txtUser Is Nothing Then Exit Sub
    
    Dim strLogin As String
    'function Login(strUserName,strPassword)
    Dim Purl As String
    Purl = strurl
    Purl = Replace(Purl, "http://", "")
    Purl = Replace(Purl, "https://", "")
    Purl = Replace(Purl, "/", "")
    Purl = Replace(Purl, ":", "")
     Dim fobj As New FileSystemObject
    Dim stmText As TextStream
    Dim strJson As String
    Set stmText = fobj.OpenTextFile(App.Path & "\" & Purl & "_pure.js")
    strJscript = stmText.ReadAll()
    strLogin = "{0};Login('{1}','{2}')"
    strJsAll = Replace(strLogin, "{0}", strJscript)
    strJsAll = Replace(strJsAll, "{1}", Me.Gobalobj.ClientUserName)
    strJsAll = Replace(strJsAll, "{2}", Me.Gobalobj.ClientPassword)
    'If InStr(1, strurl, "kcai") > 0 Then
        doc.parentWindow.execScript strJsAll
    'End If
    'If InStr(1, strurl, "mg") > 0 Then
        WebBrowser1.Silent = True
    'End If
    Gobalobj.ExchangeSwitched = True
    Exit Sub
    
    Dim objname As IHTMLElement
    Dim objpwd As IHTMLElement
    Dim objbtn As IHTMLElement
    Set objname = doc.getElementById("txt_username")
    Set objpwd = doc.getElementById("txt_pwd")
    Set objbtn = doc.getElementById("login-submit-Button")
    If objname Is Nothing Or objpwd Is Nothing Then Exit Sub
    objname.setAttribute "value", gobj.ClientUserName
    objpwd.setAttribute "value", gobj.ClientPassword
    'objbtn.onclick
End Sub

Public Function RealSendMsg(Expect As String, msg As String, cnt As Integer, Optional ByRef amt As Currency = 0) As Boolean
     On Error Resume Next
     If doc Is Nothing Then
        Exit Function
     End If
         buffScript = ""
     Dim fobj As New FileSystemObject
    Dim stmText As TextStream
    Dim url As String
    url = strurl
    url = Replace(url, "http://", "")
    url = Replace(url, "https://", "")
    url = Replace(url, ":", "")
    url = Replace(url, "/", "")
    Set stmText = fobj.OpenTextFile(App.Path & "\" & url & "_pure.js")
    strJscript = stmText.ReadAll()
    stmText.Close
    Dim c2i As New CCS2InstrClass
    c2i.cJsOdds = Me.Gobalobj.Odds
    Dim strJson As String
    If Me.Gobalobj.LoginInstFillOrEnCode = 1 Then '填充模式
        strJson = cm.ToSerial(msg, Me.Gobalobj.MinChips)
    Else '编码模式
        strJson = c2i.InstrToJsonString(msg) 'Trim(Me.parentForm.cm.ToSerial(msg))
    End If
    Dim strJsAll As String
    strJsAll = Replace(Replace(Replace("{0};SendMsg('{1}','{2}')", "{0}", strJscript), "{1}", Expect), "{2}", Trim(strJson))
    'Me.WebBrowser1.RegisterAsBrowser = True
     
     'Set doc = Me.WebBrowser1.Document
    
     'MsgBox doc.body.innerHTML
    'Me.WebBrowser1.Document.body.innerHTML ' doc.parentWindow.Parent.Document.body.innerHTML
     'Me.gobj.LogObj.Log doc.body.innerHTML
     If gobj.LoginInFrame Then
        If doc.parentWindow.frames("mainFrame") Is Nothing Then
        Else
          Set doc = doc.parentWindow.frames("mainFrame").Document
        End If
     End If
         buffScript = strJsAll
     Me.WebBrowser1.Silent = True
     doc.parentWindow.execScript buffScript
     'Me.WebBrowser1.Silent = False
     Dim objPoint As IHTMLElement
     Set objPoint = doc.getElementById("userGamePointId")
     If objPoint Is Nothing Then Exit Function
     amt = objPoint.getAttribute("data")
     Me.Show
     
End Function


