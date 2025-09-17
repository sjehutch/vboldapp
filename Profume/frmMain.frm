VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   Caption         =   "ProFume"
   ClientHeight    =   10425
   ClientLeft      =   390
   ClientTop       =   630
   ClientWidth     =   14685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   14685
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMonitorStatus 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   103
      Top             =   1080
      Width           =   14835
      Begin VB.CommandButton cmdValidDateStatus 
         Height          =   855
         Left            =   12720
         TabIndex        =   142
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdNextStep7 
         Caption         =   "Next Step >>>"
         Height          =   375
         Left            =   13080
         TabIndex        =   140
         Top             =   8400
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpEndExposureTime2 
         Height          =   375
         Left            =   10440
         TabIndex        =   104
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "h:mm tt"
         Format          =   20447235
         UpDown          =   -1  'True
         CurrentDate     =   37524
      End
      Begin MSComCtl2.DTPicker dtpEndExposureDate2 
         Height          =   375
         Left            =   10440
         TabIndex        =   105
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   20447235
         CurrentDate     =   37524
      End
      Begin MSComCtl2.DTPicker dtpStartExposureTime2 
         Height          =   375
         Left            =   6000
         TabIndex        =   106
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "h:mm tt"
         Format          =   20447235
         UpDown          =   -1  'True
         CurrentDate     =   37524
      End
      Begin MSComCtl2.DTPicker dtpStartExposureDate2 
         Height          =   375
         Left            =   6000
         TabIndex        =   107
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   20447235
         CurrentDate     =   37524
      End
      Begin MSComCtl2.DTPicker dtpStartIntroTime2 
         Height          =   375
         Left            =   1920
         TabIndex        =   108
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "h:mm tt"
         Format          =   20447235
         UpDown          =   -1  'True
         CurrentDate     =   37524
      End
      Begin MSComCtl2.DTPicker dtpStartIntroDate2 
         Height          =   375
         Left            =   1920
         TabIndex        =   109
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   20447235
         CurrentDate     =   37524
      End
      Begin VSFlex7LCtl.VSFlexGrid grdMonitorStatus 
         Height          =   6135
         Left            =   0
         TabIndex        =   110
         Top             =   1800
         Width           =   14175
         _cx             =   25003
         _cy             =   10821
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label lblEETime2 
         Caption         =   "lblEETime2"
         Height          =   255
         Left            =   9000
         TabIndex        =   131
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblEEDate2 
         Caption         =   "lblEEDate2"
         Height          =   255
         Left            =   9000
         TabIndex        =   130
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblSETime2 
         Caption         =   "lblSETime2"
         Height          =   255
         Left            =   3960
         TabIndex        =   129
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblSEDate2 
         Caption         =   "lblSEDate2"
         Height          =   255
         Left            =   3960
         TabIndex        =   128
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblSITime2 
         Caption         =   "lblSITime2"
         Height          =   255
         Left            =   480
         TabIndex        =   127
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblSIDate2 
         Caption         =   "lblSIDate2"
         Height          =   255
         Left            =   480
         TabIndex        =   126
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblStartIntro2 
         Caption         =   "lblStartIntro2"
         Height          =   255
         Left            =   480
         TabIndex        =   113
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label lblStartExpo2 
         Caption         =   "lblStartExpo2"
         Height          =   255
         Left            =   3960
         TabIndex        =   112
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lblEndExpo2 
         Caption         =   "lblEndExpo2"
         Height          =   255
         Left            =   9000
         TabIndex        =   111
         Top             =   0
         Width           =   1455
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   15360
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.Frame fraDosagePlan 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   46
      Top             =   960
      Width           =   14835
      Begin VB.Frame fraGeneralInfo 
         Caption         =   "General Info (Optional)"
         Height          =   975
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   14535
         Begin VB.TextBox txtSiteName 
            Height          =   375
            Left            =   960
            MaxLength       =   50
            TabIndex        =   47
            Text            =   "txtSiteName"
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtJobName 
            Height          =   375
            Left            =   4200
            MaxLength       =   50
            TabIndex        =   48
            Text            =   "txtJobName"
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtFumigationDate 
            Height          =   375
            Left            =   7920
            MaxLength       =   50
            TabIndex        =   49
            Text            =   "txtFumigationDate"
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtLicFumigator 
            Height          =   375
            Left            =   12000
            MaxLength       =   50
            TabIndex        =   50
            Text            =   "txtLicFumigator"
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblSiteName 
            Caption         =   "Site Name:"
            Height          =   375
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblJobName 
            Caption         =   "Job Name:"
            Height          =   375
            Left            =   3360
            TabIndex        =   73
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblFumigationDate 
            Caption         =   "Fumigation Date:"
            Height          =   495
            Left            =   6600
            TabIndex        =   72
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblLicFumigator 
            Caption         =   "Licensed Fumigator:"
            Height          =   375
            Left            =   10440
            TabIndex        =   71
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraAreaInfo 
         Caption         =   "Structure/Area Info (A structure may have more than one Area)"
         Height          =   3495
         Left            =   0
         TabIndex        =   62
         Top             =   3000
         Width           =   14535
         Begin VB.CommandButton cmdCalculate 
            Caption         =   "Calculate"
            Height          =   375
            Left            =   12720
            TabIndex        =   58
            Top             =   2880
            Width           =   1695
         End
         Begin VSFlex7LCtl.VSFlexGrid grdAreaInfo 
            Height          =   2895
            Left            =   360
            TabIndex        =   57
            Top             =   360
            Width           =   12255
            _cx             =   21616
            _cy             =   5106
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483640
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
      End
      Begin VB.Frame fraResults 
         Caption         =   "Results"
         Height          =   1935
         Left            =   0
         TabIndex        =   59
         Top             =   6480
         Width           =   14535
         Begin RichTextLib.RichTextBox rtfResults 
            Height          =   1575
            Left            =   360
            TabIndex        =   60
            Top             =   240
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   2778
            _Version        =   393217
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMain.frx":0E42
         End
      End
      Begin VB.Frame fraTargetInfo 
         Caption         =   "Target Info"
         Height          =   2055
         Left            =   0
         TabIndex        =   63
         Top             =   960
         Width           =   14535
         Begin VB.CheckBox chkNoAdjSorp 
            Caption         =   "Do not compute for HLT Adjustment"
            Enabled         =   0   'False
            Height          =   375
            Left            =   11160
            TabIndex        =   150
            Top             =   1440
            Width           =   3015
         End
         Begin VB.ListBox lstTargetPests 
            Height          =   1425
            Left            =   1320
            MultiSelect     =   1  'Simple
            TabIndex        =   51
            Top             =   360
            Width           =   4095
         End
         Begin VB.ComboBox cboLifeStage 
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cboFumigationType 
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   952
            Width           =   2535
         End
         Begin VB.ComboBox cboPressureType 
            Height          =   315
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1545
            Width           =   2535
         End
         Begin VB.ComboBox cboCommodity 
            Height          =   315
            Left            =   11640
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox txtLoadFactor 
            Enabled         =   0   'False
            Height          =   375
            Left            =   11640
            MaxLength       =   3
            TabIndex        =   56
            Text            =   "txtLoadFactor"
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label lblTargetPests 
            Caption         =   "*Target Pests:"
            Height          =   615
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblLifeStage 
            Alignment       =   1  'Right Justify
            Caption         =   "*Dosage:"
            Height          =   255
            Left            =   4800
            TabIndex        =   68
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblFumigationType 
            Alignment       =   1  'Right Justify
            Caption         =   "*Fumigation Type:"
            Height          =   255
            Left            =   4800
            TabIndex        =   67
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label lblPressureType 
            Alignment       =   1  'Right Justify
            Caption         =   "*Pressure Type:"
            Height          =   255
            Left            =   4800
            TabIndex        =   66
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label lblCommodity 
            Alignment       =   1  'Right Justify
            Caption         =   "*Commodity:"
            Height          =   255
            Left            =   9360
            TabIndex        =   65
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblLoadFactor 
            Alignment       =   1  'Right Justify
            Caption         =   "*Load Factor (%):"
            Height          =   255
            Left            =   9360
            TabIndex        =   64
            Top             =   960
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdNextStep2 
         Caption         =   "Next Step >>"
         Height          =   375
         Left            =   12960
         TabIndex        =   61
         Top             =   8520
         Width           =   1575
      End
   End
   Begin VB.Frame fraIntroductionPlan 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   75
      Top             =   960
      Width           =   14835
      Begin VB.OptionButton rbtPressure 
         Caption         =   "rbtPressure"
         Height          =   495
         Left            =   3960
         TabIndex        =   148
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton rbtTemperature 
         Caption         =   "rbtTemperature"
         Height          =   495
         Left            =   2400
         TabIndex        =   147
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton btnConvert 
         Cancel          =   -1  'True
         Caption         =   "btnConvert"
         Height          =   375
         Left            =   5400
         TabIndex        =   146
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCalculateIntroRate 
         Caption         =   "lblCalculate"
         Height          =   735
         Left            =   0
         TabIndex        =   80
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdNextStep3 
         Height          =   375
         Left            =   13080
         TabIndex        =   76
         Top             =   8520
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox rtfWarnings 
         Height          =   1215
         Left            =   0
         TabIndex        =   78
         Top             =   7800
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   2143
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":0EC3
      End
      Begin VSFlex7LCtl.VSFlexGrid grdIntroRates 
         Height          =   2535
         Left            =   0
         TabIndex        =   79
         Top             =   4800
         Width           =   12975
         _cx             =   22886
         _cy             =   4471
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdIntroInput 
         Height          =   2535
         Left            =   0
         TabIndex        =   81
         Top             =   960
         Width           =   14415
         _cx             =   25426
         _cy             =   4471
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label lblCylinderTempPressure 
         Caption         =   "lblCylinderTempPressure"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   145
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label lblIntroInput 
         Caption         =   "lblIntroInput"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   119
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label lblIntroRates 
         Caption         =   "lblIntroRates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   45
         Top             =   4440
         Width           =   5775
      End
      Begin VB.Label lblWarnings 
         Caption         =   "Warnings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   0
         TabIndex        =   82
         Top             =   7440
         Width           =   1935
      End
   End
   Begin VB.Frame fraIntroHistory 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   85
      Top             =   960
      Width           =   14835
      Begin VB.CommandButton cmdNextStep5 
         Height          =   375
         Left            =   13080
         TabIndex        =   125
         Top             =   8520
         Width           =   1575
      End
      Begin VSFlex7LCtl.VSFlexGrid grdFumIntro 
         Height          =   2655
         Left            =   1680
         TabIndex        =   88
         Top             =   360
         Width           =   11415
         _cx             =   20135
         _cy             =   4683
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.TextBox txtGrandTot 
         BackColor       =   &H8000000B&
         Height          =   375
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   6480
         Width           =   2415
      End
      Begin VSFlex7LCtl.VSFlexGrid grdAreaSubTot 
         Height          =   2655
         Left            =   1680
         TabIndex        =   87
         Top             =   3600
         Width           =   5775
         _cx             =   10186
         _cy             =   4683
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label lblGrandTot 
         Caption         =   "lblGrandTot"
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label lblSubTot 
         Caption         =   "lblSubTot"
         Height          =   495
         Left            =   120
         TabIndex        =   123
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblFumIntro 
         Caption         =   "lblFumIntro"
         Height          =   615
         Left            =   120
         TabIndex        =   122
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame fraGraph 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   114
      Top             =   960
      Width           =   14835
      Begin VB.Frame fraFilterBy 
         Height          =   3735
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Width           =   2055
         Begin VB.ListBox lstDataFilter2 
            Height          =   3375
            ItemData        =   "frmMain.frx":0F49
            Left            =   120
            List            =   "frmMain.frx":0F4B
            MultiSelect     =   2  'Extended
            TabIndex        =   137
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtSelectedN 
         Height          =   285
         Left            =   0
         TabIndex        =   139
         Top             =   4680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtSelectedY 
         Height          =   285
         Left            =   0
         TabIndex        =   138
         Top             =   4320
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdExportGraph 
         Caption         =   "cmdExportGraph"
         Height          =   375
         Left            =   12960
         TabIndex        =   135
         Top             =   8400
         Width           =   1575
      End
      Begin VB.CommandButton cmdSetProperties 
         Caption         =   "cmdSetProperties"
         Height          =   375
         Left            =   10560
         TabIndex        =   134
         Top             =   8400
         Width           =   2295
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "cmdOptions"
         Height          =   375
         Left            =   8880
         TabIndex        =   133
         Top             =   8400
         Width           =   1575
      End
      Begin ChartfxLibCtl.ChartFX ChartFX1 
         Height          =   8175
         Left            =   2160
         TabIndex        =   132
         Top             =   120
         Width           =   12495
         _cx             =   22040
         _cy             =   14420
         Build           =   21
         TypeMask        =   109577217
         Style           =   -135397381
         LeftGap         =   55
         BottomGap       =   61
         DblClk          =   2
         MarkerShape     =   1
         BorderWidth     =   2
         BorderColor     =   16777230
         AxesStyle       =   2
         Axis(0).Max     =   90
         Axis(0).Decimals=   0
         Axis(0).TickMark=   -32765
         Axis(0).GridColor=   16777252
         Axis(2).Min     =   1
         Axis(2).Max     =   10
         Axis(2).Style   =   10280
         Axis(2).TickMark=   -32767
         Axis(2).Format  =   1
         Axis(2).GridColor=   16777252
         Axis(2).Format  =   1
         RGBBk           =   16777216
         RGB2DBk         =   16777216
         RGB3DBk         =   16777215
         nColors         =   16
         Axis(2).FontColor=   16777215
         Axis(0).FontColor=   16777215
         LegendFontMask  =   268435464
         BeginProperty LegendFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PointLabelsFontMask=   268435464
         BeginProperty PointLabelsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         nPts            =   1
         nSer            =   100
         NumPoint        =   1
         NumSer          =   100
         MMask           =   768
         Title(0)        =   "Concentration"
         Title(2)        =   "Test Chart"
         Title(3)        =   "Time"
         BorderS         =   8
         _Data_          =   "frmMain.frx":0F4D
      End
   End
   Begin VB.Frame fraMonitorDataInput 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   89
      Top             =   960
      Width           =   14835
      Begin VB.CommandButton cmdValidDate 
         Height          =   855
         Left            =   12720
         TabIndex        =   141
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdNextStep6 
         Caption         =   "Next"
         Height          =   375
         Left            =   13080
         TabIndex        =   9
         Top             =   8520
         Width           =   1575
      End
      Begin VB.ListBox lstDataFilter 
         Height          =   3960
         Left            =   360
         MultiSelect     =   2  'Extended
         TabIndex        =   7
         Top             =   3000
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpEndExposureDate 
         Height          =   375
         Left            =   10440
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   20447235
         CurrentDate     =   37551
      End
      Begin MSComCtl2.DTPicker dtpStartExposureDate 
         Height          =   375
         Left            =   6000
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   20447235
         CurrentDate     =   37551
      End
      Begin MSComCtl2.DTPicker dtpStartIntroDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   20447235
         CurrentDate     =   37551
      End
      Begin MSComCtl2.DTPicker dtpDummyTime 
         Height          =   255
         Left            =   5760
         TabIndex        =   90
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "h:mm tt"
         Format          =   20447235
         UpDown          =   -1  'True
         CurrentDate     =   37551
      End
      Begin MSComCtl2.DTPicker dtpDummyDate 
         Height          =   255
         Left            =   4080
         TabIndex        =   91
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   20447235
         CurrentDate     =   37551
      End
      Begin VSFlex7LCtl.VSFlexGrid grdMonitorDataInput 
         Height          =   6015
         Left            =   3000
         TabIndex        =   8
         Top             =   2280
         Width           =   11655
         _cx             =   20558
         _cy             =   10610
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin MSComCtl2.DTPicker dtpStartIntroTime 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "h:mm tt"
         Format          =   20447235
         UpDown          =   -1  'True
         CurrentDate     =   37523
      End
      Begin MSComCtl2.DTPicker dtpStartExposureTime 
         Height          =   375
         Left            =   6000
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "h:mm tt"
         Format          =   20447235
         UpDown          =   -1  'True
         CurrentDate     =   37523
      End
      Begin MSComCtl2.DTPicker dtpEndExposureTime 
         Height          =   375
         Left            =   10440
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "h:mm tt"
         Format          =   20447235
         UpDown          =   -1  'True
         CurrentDate     =   37523
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   15360
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblMonDataInput 
         Caption         =   "lblMonDataInput"
         Height          =   255
         Left            =   360
         TabIndex        =   102
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblDataFilter 
         Caption         =   "lblDataFilter"
         Height          =   255
         Left            =   360
         TabIndex        =   101
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label lblEETime 
         Caption         =   "lblEETime"
         Height          =   255
         Left            =   9000
         TabIndex        =   100
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label LblEEDate 
         Caption         =   "LblEEDate"
         Height          =   255
         Left            =   9000
         TabIndex        =   99
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblSETime 
         Caption         =   "lblSETime"
         Height          =   255
         Left            =   4560
         TabIndex        =   98
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSEDate 
         Caption         =   "lblSEDate"
         Height          =   255
         Left            =   4560
         TabIndex        =   97
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblSITime 
         Caption         =   "lblSITime"
         Height          =   255
         Left            =   480
         TabIndex        =   96
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblSIDate 
         Caption         =   "lblSIDate"
         Height          =   255
         Left            =   480
         TabIndex        =   95
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblEndExpo 
         Caption         =   "lblEndExpo"
         Height          =   255
         Left            =   9000
         TabIndex        =   94
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblStartExpo 
         Caption         =   "lblStartExpo"
         Height          =   255
         Left            =   4560
         TabIndex        =   93
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblStartIntro 
         Caption         =   "lblStartIntro"
         Height          =   255
         Left            =   480
         TabIndex        =   92
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame fraSplash 
      Height          =   9615
      Left            =   120
      TabIndex        =   115
      Top             =   480
      Width           =   15195
      Begin VB.Frame Frame1 
         Height          =   4455
         Left            =   2160
         TabIndex        =   116
         Top             =   4320
         Width           =   10335
         Begin VB.Label lblGetStarted 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4095
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   10095
         End
         Begin VB.Label lblGetHeading 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   117
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Image imaLogo 
         Height          =   645
         Left            =   5160
         Picture         =   "frmMain.frx":105C
         Top             =   960
         Width           =   3750
      End
      Begin VB.Label lblTradeMark 
         Caption         =   "lblTradeMark"
         Height          =   255
         Left            =   1920
         TabIndex        =   144
         Top             =   3840
         Width           =   10335
      End
      Begin VB.Label lblCopyright 
         Caption         =   "lblCopyright"
         Height          =   255
         Left            =   1920
         TabIndex        =   143
         Top             =   4080
         Width           =   6135
      End
      Begin VB.Label lblAboutProfume 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   1920
         TabIndex        =   118
         Top             =   1800
         Width           =   10815
      End
   End
   Begin VB.Frame fraMonitorPlan 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   120
      TabIndex        =   83
      Top             =   960
      Width           =   14835
      Begin VB.CommandButton cmdNextStep4 
         Caption         =   "Next Step >>>"
         Height          =   375
         Left            =   13080
         TabIndex        =   121
         Top             =   8520
         Width           =   1575
      End
      Begin VSFlex7LCtl.VSFlexGrid grdMonPlan 
         Height          =   3915
         Left            =   0
         TabIndex        =   84
         Top             =   360
         Width           =   12915
         _cx             =   22781
         _cy             =   6906
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label lblMonitorPlan 
         Caption         =   "lblMonitorPlan"
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   0
         Width           =   9615
      End
   End
   Begin VB.Frame fraCustomerInfo 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   14835
      Begin VB.Frame fraAddInfo 
         Caption         =   "Address Information"
         Height          =   2055
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   14535
         Begin VB.TextBox txtCompany 
            Height          =   285
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   0
            Top             =   360
            Width           =   4575
         End
         Begin VB.TextBox txtAddress1 
            Height          =   285
            Left            =   1800
            MaxLength       =   150
            TabIndex        =   10
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox txtAddress2 
            Height          =   285
            Left            =   1800
            MaxLength       =   150
            TabIndex        =   11
            Top             =   1080
            Width           =   4575
         End
         Begin VB.TextBox txtAddress3 
            Height          =   285
            Left            =   1800
            MaxLength       =   150
            TabIndex        =   12
            Top             =   1440
            Width           =   4575
         End
         Begin VB.TextBox txtCity 
            Height          =   285
            Left            =   9840
            MaxLength       =   50
            TabIndex        =   14
            Top             =   360
            Width           =   4575
         End
         Begin VB.ComboBox cboState 
            Height          =   315
            Left            =   9840
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtZip 
            Height          =   285
            Left            =   9840
            MaxLength       =   10
            TabIndex        =   16
            Top             =   1080
            Width           =   2295
         End
         Begin VB.ComboBox cboCountry 
            Height          =   315
            Left            =   9840
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label lblCompany 
            Caption         =   "Company"
            Height          =   375
            Left            =   360
            TabIndex        =   44
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblAddress1 
            Caption         =   "Address Line 1:"
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblAddress2 
            Caption         =   "Address Line 2:"
            Height          =   375
            Left            =   360
            TabIndex        =   42
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblAddress3 
            Caption         =   "Address Line 3:"
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblCity 
            Caption         =   "City:"
            Height          =   255
            Left            =   8280
            TabIndex        =   40
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblState 
            Caption         =   "State/Province:"
            Height          =   255
            Left            =   8280
            TabIndex        =   39
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblZip 
            Caption         =   "Zip/Postal Code:"
            Height          =   255
            Left            =   8280
            TabIndex        =   38
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label lblCountry 
            Caption         =   "Country:"
            Height          =   255
            Left            =   8280
            TabIndex        =   37
            Top             =   1440
            Width           =   855
         End
      End
      Begin VB.Frame fraNotes 
         Caption         =   "Notes"
         Height          =   3975
         Left            =   120
         TabIndex        =   28
         Top             =   4440
         Width           =   14535
         Begin VB.TextBox txtNotes 
            Height          =   3375
            Left            =   1800
            MaxLength       =   10000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   360
            Width           =   12375
         End
         Begin VB.Label lblNotes 
            Caption         =   "Notes"
            Height          =   255
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraContactInfo 
         Caption         =   "Contact Information"
         Height          =   2055
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Width           =   14535
         Begin VB.TextBox txtContactName1 
            Height          =   285
            Left            =   1800
            MaxLength       =   80
            TabIndex        =   18
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtContactPhone1 
            Height          =   285
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   19
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtContactEmail1 
            Height          =   285
            Left            =   1800
            MaxLength       =   60
            TabIndex        =   20
            Top             =   1440
            Width           =   1935
         End
         Begin VB.TextBox txtContactName2 
            Height          =   285
            Left            =   4680
            MaxLength       =   80
            TabIndex        =   21
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtContactPhone2 
            Height          =   285
            Left            =   4680
            MaxLength       =   20
            TabIndex        =   22
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtContactEmail2 
            Height          =   285
            Left            =   4680
            MaxLength       =   60
            TabIndex        =   23
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblPrimary 
            Caption         =   "Primary Contact Info:"
            Height          =   255
            Left            =   1800
            TabIndex        =   35
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblName 
            Caption         =   "Name:"
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblPhone 
            Caption         =   "Phone:"
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblEmail 
            Caption         =   "E-Mail:"
            Height          =   255
            Left            =   360
            TabIndex        =   32
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblSecondary 
            Caption         =   "Secondary Contact Info:"
            Height          =   255
            Left            =   4680
            TabIndex        =   31
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdNextStep1 
         Caption         =   "Next Step >>"
         Height          =   375
         Left            =   13080
         TabIndex        =   25
         Top             =   8520
         Width           =   1575
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   10155
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   17701
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   149
      Top             =   0
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A10
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B22
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C34
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D46
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E58
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F6A
            Key             =   "Paste"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabTabs 
      Height          =   9735
      Left            =   0
      TabIndex        =   26
      Top             =   360
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   17171
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Customer Info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Fumigation Dosage Plan"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Introduction Plan"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Monitoring Plan"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Introduction History"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Monitoring Data Input"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Monitoring Status"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Graph"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Begin VB.Menu mnuFileNew 
         Caption         =   "1001"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1002"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "1003"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "1004"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "1005"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "1018"
      End
      Begin VB.Menu mnuFileExportGraph 
         Caption         =   "1019"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "1006"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "1007"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "1008"
      Begin VB.Menu mnuEditCut 
         Caption         =   "1018"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "1009"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "1010"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1011"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1012"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1013"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "1014"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1015"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "1016"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1017"
      End
   End
   Begin VB.Menu mnuDosagePlanContext 
      Caption         =   "Dosage Plan Context"
      Visible         =   0   'False
      Begin VB.Menu mnuDosagePlanDelete 
         Caption         =   "10100"
      End
   End
   Begin VB.Menu mnuIntroPlanContext 
      Caption         =   "Intro Plan Context"
      Visible         =   0   'False
      Begin VB.Menu mnuIntroPlanDelete 
         Caption         =   "10100"
      End
   End
   Begin VB.Menu mnuMonitorPlanContext 
      Caption         =   "Monitor Plan Context"
      Visible         =   0   'False
      Begin VB.Menu mnuMonitorPlanDelete 
         Caption         =   "10100"
      End
   End
   Begin VB.Menu mnuIntroHistoryContext 
      Caption         =   "Intro History Context"
      Visible         =   0   'False
      Begin VB.Menu mnuIntroHistoryDelete 
         Caption         =   "10100"
      End
   End
   Begin VB.Menu mnuMonitorInputContext 
      Caption         =   "Monitor Input Context"
      Visible         =   0   'False
      Begin VB.Menu mnuMonitorInputCopy 
         Caption         =   "10092"
      End
      Begin VB.Menu mnuMonitorInputPaste 
         Caption         =   "10093"
      End
      Begin VB.Menu mnuMonitorInputDelete 
         Caption         =   "10100"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const MODULE_NAME = "frmMain.frm"
Dim Row, Col

Private Sub btnConvert_Click()
   Dim I
    gblPresTempMode = True
    Call ConvertPresTemp
    frmMain.cmdCalculateIntroRate.Enabled = True
    frmMain.btnConvert.Enabled = False
        frmMain.grdIntroInput.Enabled = True
        For I = 1 To 9
            frmMain.grdIntroInput.Cell(flexcpForeColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H0
            frmMain.grdIntroInput.Cell(flexcpBackColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H80000005
        Next
End Sub

Private Sub cboCommodity_Click()
    cboCommodity_Click_Handler
    Call EnableAdjHLTSorptionChkBox
End Sub


Private Sub cboCommodity_Validate(Cancel As Boolean)
    cboCommodity_Validate_Handler Cancel
End Sub

Private Sub cboCountry_Click()
    cboCountry_Click_Handler
End Sub


Private Sub cboCountry_Validate(Cancel As Boolean)
    cboCountry_Validate_Handler Cancel
End Sub

Private Sub cboFumigationType_Click()
    cboFumigationType_Click_Handler
End Sub


Private Sub cboFumigationType_Validate(Cancel As Boolean)
    cboFumigationType_Validate_Handler Cancel
End Sub

Private Sub cboLifeStage_Click()
    cboLifeStage_Click_Handler
End Sub
'Pat 2/10/06 - Added, A warning message should appear if the user check the checkbox
Private Sub chkNoAdjSorp_Click()
    If frmMain.chkNoAdjSorp.Value = 1 Then
        MsgBox LoadResString(IDS_ADJUSTED_SORP_MSGBOX), vbCritical
    End If
End Sub

Private Sub chkNoAdjSorp_Validate(Cancel As Boolean)
    chkNoAdjSorp_Validate_Handler Cancel
End Sub

Private Sub cboLifeStage_Validate(Cancel As Boolean)
    cboLifeStage_Validate_Handler Cancel
End Sub

Private Sub cboPressureType_Click()
    cboPressureType_Click_Handler
End Sub


Private Sub cboPressureType_Validate(Cancel As Boolean)
    cboPressureType_Validate_Handler Cancel
End Sub

Private Sub cboState_Click()
    cboState_Click_Handler
End Sub


Private Sub cboState_Validate(Cancel As Boolean)
    cboState_Validate_Handler Cancel
End Sub



Private Sub cmdCalculate_Click()
    cmdCalculate_Click_Handler
End Sub
Private Sub cmdCalculateIntroRate_Click()
    cmdCalculateIntroRate_Click_Handler
End Sub

Private Sub cmdExportGraph_Click()
    cmdExportGraph_Click_Handler
End Sub

Private Sub cmdNextStep1_Click()
    NextTab
End Sub

Private Sub cmdNextStep2_Click()
    NextTab
End Sub

Private Sub cmdNextStep3_Click()
    NextTab
End Sub

Private Sub cmdNextStep4_Click()
    NextTab
End Sub

Private Sub cmdNextStep5_Click()
    NextTab
End Sub


Private Sub cmdNextStep6_Click()
    NextTab
End Sub

Private Sub cmdNextStep7_Click()
    NextTab
End Sub

Private Sub cmdOptions_Click()
    cmdOptions_Click_Handler
End Sub

Private Sub cmdSetProperties_Click()
    cmdSetProperties_Click_Handler
End Sub

Private Sub cmdValidDate_Click()
    cmdValidDate_Click_Handler
End Sub

Private Sub cmdValidDateStatus_Click()
    cmdValidDateStatus_Click_Handler
End Sub

Private Sub dtpDummyDate_KeyDown(KeyCode As Integer, Shift As Integer)
    dtpDummyDate_KeyDown_Handler KeyCode, Shift
End Sub

Private Sub dtpDummyDate_LostFocus()
    dtpDummyDate_LostFocus_Handler
End Sub


Private Sub dtpDummyTime_KeyDown(KeyCode As Integer, Shift As Integer)
    dtpDummyTime_KeyDown_Handler KeyCode, Shift
End Sub

Private Sub dtpDummyTime_LostFocus()
    dtpDummyTime_LostFocus_Handler
End Sub

Private Sub dtpEndExposureDate_Change()
    dtpEndExposureDate_Change_Handler
End Sub

Private Sub dtpEndExposureDate2_Change()
    dtpEndExposureDate2_Change_Handler
End Sub




Private Sub dtpEndExposureTime_Change()
    dtpEndExposureTime_Change_Handler
End Sub

Private Sub dtpEndExposureTime2_Change()
    dtpEndExposureTime2_Change_Handler
End Sub





Private Sub dtpStartExposureDate_Change()
    dtpStartExposureDate_Change_Handler
End Sub

Private Sub dtpStartExposureDate2_Change()
    dtpStartExposureDate2_Change_Handler
End Sub









Private Sub dtpStartExposureTime_Change()
    dtpStartExposureTime_Change_Handler
End Sub

Private Sub dtpStartExposureTime2_Change()
    dtpStartExposureTime2_Change_Handler
End Sub





Private Sub dtpStartIntroDate_Change()
    dtpStartIntroDate_Change_Handler
End Sub

Private Sub dtpStartIntroDate2_Change()
    dtpStartIntroDate2_Change_Handler
End Sub

Private Sub dtpStartIntroTime_Change()
    dtpStartIntroTime_Change_Handler
End Sub

Private Sub dtpStartIntroTime2_Change()
    dtpStartIntroTime2_Change_Handler
End Sub

Private Sub Form_Load()
    Load_Main
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Form_QueryUnload_Handler Cancel, UnloadMode
End Sub





Private Sub grdAreaInfo_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    grdAreaInfo_AfterEdit_Handler Row, Col
End Sub

Private Sub grdAreaInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdAreaInfo_BeforeUserResize_Handler Row, Col, Cancel
End Sub

Private Sub grdAreaInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    grdAreaInfo_KeyDown_Handler KeyCode, Shift
End Sub

Private Sub grdAreaInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdAreaInfo_MouseDown_Handler Button, Shift, x, Y
End Sub

Private Sub grdAreaInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdAreaInfo_MouseMove_Handler Button, Shift, x, Y
End Sub

Private Sub grdAreaInfo_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdAreaInfo_StartEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdAreaInfo_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdAreaInfo_ValidateEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdAreaSubTot_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    grdAreaSubTot_AfterUserResize_Handler Row, Col
End Sub

Private Sub grdAreaSubTot_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdAreaSubTot_BeforeEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdFumIntro_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    grdFumIntro_AfterEdit_Handler Row, Col
End Sub

Private Sub grdFumIntro_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    grdFumIntro_AfterUserResize_Handler Row, Col
End Sub

Private Sub grdFumIntro_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdFumIntro_BeforeEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdFumIntro_KeyDown(KeyCode As Integer, Shift As Integer)
    grdFumIntro_KeyDown_Handler KeyCode, Shift
End Sub

Private Sub grdFumIntro_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdFumIntro_MouseDown_Handler Button, Shift, x, Y
End Sub

Private Sub grdFumIntro_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdFumIntro_MouseMove_Handler Button, Shift, x, Y
End Sub

Private Sub grdFumIntro_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdFumIntro_StartEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdFumIntro_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdFumIntro_ValidateEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdIntroInput_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    grdIntroInput_AfterEdit_Handler Row, Col
End Sub

Private Sub grdIntroInput_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdIntroInput_BeforeUserResize_Handler Row, Col, Cancel
End Sub

Private Sub grdIntroInput_KeyDown(KeyCode As Integer, Shift As Integer)
    grdIntroInput_KeyDown_Handler KeyCode, Shift
End Sub

Private Sub grdIntroInput_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdIntroInput_MouseDown_Handler Button, Shift, x, Y
End Sub

Private Sub grdIntroInput_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdIntroInput_MouseMove_Handler Button, Shift, x, Y
End Sub

Private Sub grdIntroInput_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdIntroInput_StartEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdIntroInput_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdIntroInput_ValidateEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdIntroRates_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdIntroRates_BeforeEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdIntroRates_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdIntroRates_MouseMove_Handler Button, Shift, x, Y
End Sub

Private Sub grdMonitorDataInput_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    grdMonitorDataInput_AfterEdit_Handler Row, Col
End Sub

Private Sub grdMonitorDataInput_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single, Cancel As Boolean)
    grdMonitorDataInput_BeforeMouseDown_Handler Button, Shift, x, Y, Cancel
End Sub

Private Sub grdMonitorDataInput_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    grdMonitorDataInput_BeforeScroll_Handler OldTopRow, OldLeftCol, NewTopRow, NewLeftCol, Cancel
End Sub

Private Sub grdMonitorDataInput_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonitorDataInput_BeforeUserResize_Handler Row, Col, Cancel
End Sub


Private Sub grdMonitorDataInput_KeyDown(KeyCode As Integer, Shift As Integer)
    grdMonitorDataInput_KeyDown_Handler KeyCode, Shift
End Sub

Private Sub grdMonitorDataInput_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdMonitorDataInput_MouseDown_Handler Button, Shift, x, Y
End Sub

Private Sub grdMonitorDataInput_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdMonitorDataInput_MouseMove_Handler Button, Shift, x, Y
End Sub

Private Sub grdMonitorDataInput_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonitorDataInput_StartEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdMonitorDataInput_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonitorDataInput_ValidateEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdMonitorStatus_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonitorStatus_BeforeUserResize_Handler Row, Col, Cancel
End Sub


Private Sub grdMonitorStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdMonitorStatus_MouseMove_Handler Button, Shift, x, Y
End Sub

Private Sub grdMonPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    grdMonPlan_AfterEdit_Handler Row, Col
End Sub

Private Sub grdMonPlan_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    grdMonPlan_AfterUserResize_Handler Row, Col
End Sub

Private Sub grdMonPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonPlan_BeforeEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdMonPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonPlan_BeforeUserResize_Handler Row, Col, Cancel
End Sub

Private Sub grdMonPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    grdMonPlan_KeyDown_Handler KeyCode, Shift
End Sub

Private Sub grdMonPlan_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdMonPlan_MouseDown_Handler Button, Shift, x, Y
End Sub

Private Sub grdMonPlan_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    grdMonPlan_MouseMove_Handler Button, Shift, x, Y
End Sub

Private Sub grdMonPlan_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonPlan_StartEdit_Handler Row, Col, Cancel
End Sub

Private Sub grdMonPlan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    grdMonPlan_ValidateEdit_Handler Row, Col, Cancel
End Sub


Private Sub lstDataFilter_Click()
    lstDataFilter_Click_Handler
End Sub


Private Sub lstDataFilter2_Click()
    lstDataFilter2_Click_Handler
End Sub

Private Sub lstTargetPests_GotFocus()
    lstTargetPests_GotFocus_Handler
End Sub

Private Sub lstTargetPests_Validate(Cancel As Boolean)
    lstTargetPests_Validate_Handler Cancel
End Sub

Private Sub mnuDosagePlanDelete_Click()
    mnuDosagePlanDelete_Click_Handler
End Sub

Private Sub mnuEditCopy_Click()
    mnuEditCopy_Click_Handler
End Sub

Private Sub mnuEditCut_Click()
    mnuEditCut_Click_Handler
End Sub

Private Sub mnuEditPaste_Click()
    mnuEditPaste_Click_Handler
End Sub
Private Sub mnuFileClose_Click()
    CloseFile
End Sub
Private Sub mnuFileExport_Click()
    ExportFile
End Sub

Private Sub mnuFileExportGraph_Click()
    'JNM 04/27/2003 - AT Fix - Call the handler in modmain
    mnuFileExportGraph_Click_Handler
End Sub

Private Sub mnuFileNew_Click()
    NewFile
End Sub


Private Sub mnuFileOpen_Click()
    OpenFile
End Sub

Private Sub mnuFileSave_Click()
    SaveFile
End Sub


Private Sub mnuFileSaveAs_Click()
  SaveFile True
End Sub


Private Sub mnuHelpAbout_Click()
    mnuHelpAbout_Click_Handler
End Sub


Private Sub mnuHelpContents_Click()
    mnuHelpContents_Click_Handler
End Sub

Private Sub mnuIntroHistoryDelete_Click()
    mnuIntroHistoryDelete_Click_Handler
End Sub

Private Sub mnuIntroPlanDelete_Click()
    mnuIntroPlanDelete_Click_Handler
End Sub

Private Sub mnuMonitorInputCopy_Click()
    mnuMonitorInputCopy_Click_Handler
End Sub

Private Sub mnuMonitorInputDelete_Click()
    mnuMonitorInputDelete_Click_Handler
End Sub

Private Sub mnuMonitorInputPaste_Click()
    mnuMonitorInputPaste_Click_Handler
End Sub

Private Sub mnuMonitorPlanDelete_Click()
    mnuMonitorPlanDelete_Click_Handler
End Sub

Private Sub mnuViewOptions_Click()
    mnuViewOptions_Click_Handler
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar_Click_Handler
End Sub
Private Sub mnuViewToolbar_Click()
    mnuViewToolbar_Click_Handler
End Sub
Private Sub rbtPressure_Click()
    Dim I
    If blnOpenFile = True Then
        gblPresTempMode = True
        frmMain.btnConvert.Enabled = True
        blnOpenFile = False
        gblPressure = True
    End If
    blnTempPres = True
    gblPresTempMode = True
    frmMain.btnConvert.Enabled = True
    blnConvert = False
    If blnEnable = False Then
        frmMain.grdIntroInput.Enabled = True
    Else
        gbForcedClick = False
        frmMain.rbtTemperature.Enabled = False
        frmMain.grdIntroInput.Enabled = False
        For I = 1 To 9
            frmMain.grdIntroInput.Cell(flexcpForeColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H0
            frmMain.grdIntroInput.Cell(flexcpBackColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H8000000B
        Next
        frmMain.cmdCalculateIntroRate.Enabled = False
        frmMain.mnuViewOptions.Enabled = False
    End If
    
End Sub
Private Sub rbtTemperature_Click()
    Dim I
    blnTempPres = False
    gblPresTempMode = False
    frmMain.btnConvert.Enabled = True
    blnConvert = False
    If blnEnable = False Then
        frmMain.grdIntroInput.Enabled = True
    Else
        gbForcedClick = False
        frmMain.rbtPressure.Enabled = False
        frmMain.grdIntroInput.Enabled = False
        For I = 1 To 9
            frmMain.grdIntroInput.Cell(flexcpForeColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H0
            frmMain.grdIntroInput.Cell(flexcpBackColor, 1, I, 100, frmMain.grdIntroInput.Cols - 1) = &H8000000B
        Next
        frmMain.cmdCalculateIntroRate.Enabled = False
        frmMain.mnuViewOptions.Enabled = False
    End If
End Sub

Private Sub tabTabs_BeforeClick(Cancel As Integer)
    tabTabs_BeforeClick_Handler Cancel
End Sub
Private Sub tabTabs_Click()
    tabTabs_Click_Handler
End Sub
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    tbToolBar_ButtonClick_Handler Button
End Sub

Private Sub mnuFileExit_Click()
    mnuFileExit_Click_Handler
End Sub

Private Sub txtAddress1_Change()
    txtAddress1_Change_Handler
End Sub

Private Sub txtAddress1_Validate(Cancel As Boolean)
    txtAddress1_Validate_Handler Cancel
End Sub

Private Sub txtAddress2_Change()
    txtAddress2_Change_Handler
End Sub

Private Sub txtAddress2_Validate(Cancel As Boolean)
    txtAddress2_Validate_Handler Cancel
End Sub

Private Sub txtAddress3_Change()
    txtAddress3_Change_Handler
End Sub


Private Sub txtAddress3_Validate(Cancel As Boolean)
    txtAddress3_Validate_Handler Cancel
End Sub

Private Sub txtCity_Change()
    txtCity_Change_Handler
End Sub

Private Sub txtCity_Validate(Cancel As Boolean)
    txtCity_Validate_Handler Cancel
End Sub

Private Sub txtCompany_Change()
    txtCompany_Change_Handler
End Sub


Private Sub txtCompany_Validate(Cancel As Boolean)
    txtCompany_Validate_Handler Cancel
End Sub

Private Sub txtContactEmail1_Change()
    txtContactEmail1_Change_Handler
End Sub

Private Sub txtContactEmail1_Validate(Cancel As Boolean)
    txtContactEmail1_Validate_Handler Cancel
End Sub

Private Sub txtContactEmail2_Change()
    txtContactEmail2_Change_Handler
End Sub

Private Sub txtContactEmail2_Validate(Cancel As Boolean)
    txtContactEmail2_Validate_Handler Cancel
End Sub

Private Sub txtContactName1_Change()
    txtContactName1_Change_Handler
End Sub


Private Sub txtContactName1_Validate(Cancel As Boolean)
    txtContactName1_Validate_Handler Cancel
End Sub

Private Sub txtContactName2_Change()
    txtContactName2_Change_Handler
End Sub

Private Sub txtContactName2_Validate(Cancel As Boolean)
    txtContactName2_Validate_Handler Cancel
End Sub

Private Sub txtContactPhone1_Change()
    txtContactPhone1_Change_Handler
End Sub

Private Sub txtContactPhone1_Validate(Cancel As Boolean)
    txtContactPhone1_Validate_Handler Cancel
End Sub

Private Sub txtContactPhone2_Change()
    txtContactPhone2_Change_Handler
End Sub

Private Sub txtContactPhone2_Validate(Cancel As Boolean)
    txtContactPhone2_Validate_Handler Cancel
End Sub

Private Sub txtFumigationDate_Change()
    txtFumigationDate_Change_Handler
End Sub

Private Sub txtFumigationDate_Validate(Cancel As Boolean)
    txtFumigationDate_Validate_Handler Cancel
End Sub



Private Sub txtJobName_Change()
    txtJobName_Change_Handler
End Sub

Private Sub txtJobName_Validate(Cancel As Boolean)
    txtJobName_Validate_Handler Cancel
End Sub

Private Sub txtLicFumigator_Change()
    txtLicFumigator_Change_Handler
End Sub

Private Sub txtLicFumigator_Validate(Cancel As Boolean)
    txtLicFumigator_Validate_Handler Cancel
End Sub

Private Sub txtLoadFactor_Change()
    txtLoadFactor_Change_Handler
    Call EnableAdjHLTSorptionChkBox
End Sub

'Private Sub txtLoadFactor_LostFocus()
'    txtLoadFactor_LostFocus_Handler
'End Sub

Private Sub txtLoadFactor_Validate(Cancel As Boolean)
    txtLoadFactor_Validate_Handler Cancel
    
End Sub

Private Sub txtNotes_Change()
    txtNotes_Change_Handler
End Sub

Private Sub txtNotes_Validate(Cancel As Boolean)
    txtNotes_Validate_Handler Cancel
End Sub


Private Sub txtSiteName_Change()
    txtSiteName_Change_Handler
End Sub



Private Sub txtSiteName_Validate(Cancel As Boolean)
    txtSiteName_Validate_Handler Cancel
End Sub

Private Sub txtZip_Change()
    txtZip_Change_Handler
End Sub


Private Sub txtZip_Validate(Cancel As Boolean)
    txtZip_Validate_Handler Cancel
End Sub

