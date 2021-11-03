VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PromKappa 
   Caption         =   "PromKappa: Analysis of DNA sequences through Kappa Index of Coincidence method"
   ClientHeight    =   12990
   ClientLeft      =   0
   ClientTop       =   270
   ClientWidth     =   17475
   Icon            =   "CG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12990
   ScaleWidth      =   17475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame MRP 
      Caption         =   "Master results"
      Height          =   9495
      Left            =   4200
      TabIndex        =   56
      Top             =   4080
      Visible         =   0   'False
      Width           =   10335
      Begin VB.Frame Frame6 
         Caption         =   "Setting for this graph"
         Height          =   3615
         Left            =   5280
         TabIndex        =   66
         Top             =   5640
         Width           =   4935
         Begin VB.CheckBox Save_Result_MR 
            Caption         =   "Save this graph"
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Setting for this graph"
         Height          =   3615
         Left            =   120
         TabIndex        =   65
         Top             =   5640
         Width           =   4935
         Begin VB.CheckBox Show_center_PC 
            Caption         =   "Show center as points/ circles"
            Height          =   375
            Left            =   240
            TabIndex        =   74
            Top             =   720
            Width           =   3255
         End
         Begin VB.Frame Frame7 
            Caption         =   "Plot"
            Height          =   1215
            Left            =   240
            TabIndex        =   69
            Top             =   1200
            Width           =   4455
            Begin VB.OptionButton Option1 
               Caption         =   "Points"
               Height          =   255
               Left            =   240
               TabIndex        =   71
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Circles"
               Height          =   255
               Left            =   240
               TabIndex        =   70
               Top             =   720
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.CheckBox Save_Result_MP 
            Caption         =   "Save this graph"
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton Erase_Center_patt 
         Caption         =   "Erase"
         Height          =   375
         Left            =   120
         TabIndex        =   64
         Top             =   5040
         Width           =   5055
      End
      Begin VB.CommandButton Erase_MasterGraph 
         Caption         =   "Erase"
         Height          =   375
         Left            =   5160
         TabIndex        =   63
         Top             =   5040
         Width           =   5055
      End
      Begin VB.PictureBox MasterGraph 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4695
         Left            =   5160
         ScaleHeight     =   309
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   60
         Top             =   240
         Width           =   5055
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Kappa IC"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "C+G"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   4440
            TabIndex        =   61
            Top             =   4320
            Width           =   495
         End
         Begin VB.Line Line11 
            BorderStyle     =   3  'Dot
            X1              =   168
            X2              =   168
            Y1              =   0
            Y2              =   312
         End
         Begin VB.Line Line12 
            BorderStyle     =   3  'Dot
            X1              =   336
            X2              =   -8
            Y1              =   160
            Y2              =   160
         End
      End
      Begin VB.PictureBox Center_patt 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4695
         Left            =   120
         ScaleHeight     =   309
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   333
         TabIndex        =   57
         Top             =   240
         Width           =   5055
         Begin VB.Line Line9 
            BorderStyle     =   3  'Dot
            X1              =   336
            X2              =   -8
            Y1              =   160
            Y2              =   160
         End
         Begin VB.Line Line10 
            BorderStyle     =   3  'Dot
            X1              =   168
            X2              =   168
            Y1              =   0
            Y2              =   312
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "C+G"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   4440
            TabIndex        =   59
            Top             =   4320
            Width           =   495
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Kappa IC"
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   0
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Gene promoters"
      Height          =   11295
      Left            =   14640
      TabIndex        =   50
      Top             =   2280
      Width           =   2775
      Begin VB.CommandButton Search_gene_lis 
         Caption         =   "Load, search and generate"
         Height          =   375
         Left            =   120
         TabIndex        =   86
         Top             =   3240
         Width           =   2535
      End
      Begin VB.TextBox gene_search_txt 
         Height          =   285
         Left            =   1200
         TabIndex        =   85
         Text            =   "gene name"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton gene_search 
         Caption         =   "Search"
         Height          =   375
         Left            =   120
         TabIndex        =   84
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Pause_p_test 
         Caption         =   "Pause"
         Height          =   255
         Left            =   1560
         TabIndex        =   82
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Graph2_show 
         Caption         =   "Show Sequence Result"
         Height          =   615
         Left            =   1440
         TabIndex        =   73
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Graph1_show 
         Caption         =   "Show Master Result"
         Height          =   615
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Save_Result 
         Caption         =   "Save results as images (BMP)"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CommandButton Promotor 
         Caption         =   "Open promoter file"
         Height          =   495
         Left            =   120
         TabIndex        =   54
         Top             =   1080
         Width           =   2535
      End
      Begin VB.ListBox List_promotori 
         Height          =   6690
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Width           =   2535
      End
      Begin VB.CommandButton Kappa_for_promotor_list 
         Caption         =   "Test promoters"
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Stop_p_test 
         Caption         =   "Stop"
         Height          =   255
         Left            =   1560
         TabIndex        =   51
         Top             =   1920
         Width           =   1095
      End
   End
   Begin VB.Frame Colors_CHART 
      Caption         =   "KappaIC/C+G - colors"
      Height          =   1815
      Left            =   14640
      TabIndex        =   43
      Top             =   240
      Width           =   2775
      Begin VB.PictureBox Col_pic3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1560
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   49
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Col_3 
         Caption         =   "BackGround"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   1320
         Width           =   1335
      End
      Begin VB.PictureBox Col_pic2 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   47
         Top             =   840
         Width           =   1095
      End
      Begin VB.PictureBox Col_pic1 
         BackColor       =   &H00C00000&
         Height          =   375
         Left            =   1560
         ScaleHeight     =   315
         ScaleWidth      =   1035
         TabIndex        =   46
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Col_2 
         Caption         =   "to color"
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Col_1 
         Caption         =   "From color"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   14760
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Signals"
      Height          =   1815
      Left            =   11520
      TabIndex        =   24
      Top             =   2160
      Width           =   3015
      Begin VB.PictureBox CG_color_line 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   80
         Top             =   1440
         Width           =   255
      End
      Begin VB.PictureBox DI_color_line 
         BackColor       =   &H00FF0000&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   79
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox Kappa_color_line 
         BackColor       =   &H00000000&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   78
         Top             =   720
         Width           =   255
      End
      Begin VB.PictureBox TM_color_line 
         BackColor       =   &H00C00000&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   195
         TabIndex        =   77
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox PROC4 
         Caption         =   "C+G RESULTS"
         Height          =   375
         Left            =   480
         TabIndex        =   28
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox PROC3 
         Caption         =   "DINUCLEOTIDE  RESULTS"
         Height          =   375
         Left            =   480
         TabIndex        =   27
         Top             =   960
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox PROC2 
         Caption         =   "KAPPA RESULTS"
         Height          =   375
         Left            =   480
         TabIndex        =   26
         Top             =   600
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox PROC1 
         Caption         =   "TM RESULTS"
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.PictureBox graf_TM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   9360
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   16
      Top             =   8880
      Width           =   5055
      Begin VB.Line Line8 
         BorderStyle     =   3  'Dot
         X1              =   336
         X2              =   -8
         Y1              =   152
         Y2              =   152
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   312
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Kappa IC"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TM"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4440
         TabIndex        =   17
         Top             =   4320
         Width           =   495
      End
   End
   Begin VB.PictureBox graf_DI 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   9360
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   13
      Top             =   4080
      Width           =   5055
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   336
         X2              =   -8
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   312
      End
      Begin VB.Label DINUCLEO 
         BackStyle       =   0  'Transparent
         Caption         =   "CG"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Kappa IC"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox graf_IC_CG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   4200
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   10
      Top             =   8880
      Width           =   5055
      Begin VB.Line Line7 
         BorderStyle     =   3  'Dot
         X1              =   336
         X2              =   0
         Y1              =   152
         Y2              =   152
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   312
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Kappa IC"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "C+G%"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   4320
         Width           =   495
      End
   End
   Begin VB.PictureBox graf_point 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   4200
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   7
      Top             =   4080
      Width           =   5055
      Begin VB.CheckBox F_point 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Follow point"
         Height          =   255
         Left            =   3720
         TabIndex        =   81
         Top             =   0
         Width           =   1335
      End
      Begin VB.Shape Focus_Shape 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Height          =   135
         Left            =   5040
         Top             =   0
         Width           =   135
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   336
         X2              =   0
         Y1              =   160
         Y2              =   160
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   312
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "C+G%"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   4440
         TabIndex        =   9
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Kappa IC"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Results"
      Height          =   4695
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   8640
      Width           =   3855
      Begin VB.TextBox Rezultate_txt 
         Height          =   3135
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton Procesare 
         Caption         =   "Start sequence analisys"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton Stop_procesare 
         Caption         =   "Stop"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   4455
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   3855
      Begin VB.TextBox mot33 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         TabIndex        =   93
         Text            =   "0"
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox mot22 
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   3480
         TabIndex        =   92
         Text            =   "0"
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox mot11 
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3480
         TabIndex        =   91
         Text            =   "0"
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox mot3 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3120
         TabIndex        =   90
         Text            =   "0"
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox mot2 
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   3120
         TabIndex        =   89
         Text            =   "0"
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox mot1 
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3120
         TabIndex        =   88
         Text            =   "0"
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox PPDOG 
         Caption         =   "Print promoter data on graph"
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CheckBox SecSenzitiv 
         Caption         =   "Calculate the promoter nucleotide frequency"
         Height          =   375
         Left            =   240
         TabIndex        =   83
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3495
      End
      Begin VB.CommandButton Filtru3 
         Caption         =   "L"
         Height          =   375
         Left            =   720
         TabIndex        =   38
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton Filtru2 
         Caption         =   "U"
         Height          =   375
         Left            =   1200
         TabIndex        =   37
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton Filtru1 
         Caption         =   "C"
         Height          =   375
         Left            =   1680
         TabIndex        =   36
         Top             =   3840
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "DELETE RESULTS"
         Height          =   375
         Left            =   2280
         TabIndex        =   35
         Top             =   3840
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox TATA 
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1080
         TabIndex        =   34
         Text            =   "TATA"
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox ATAAA 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Text            =   "AATAAA"
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox ATG 
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   1080
         TabIndex        =   32
         Text            =   "ATG"
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton Filtru4 
         Caption         =   "G"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox Step_Window 
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.ComboBox DubleN 
         Height          =   315
         Left            =   2160
         TabIndex        =   19
         Text            =   "CG"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Lungime_fereastra 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Text            =   "30"
         Top             =   480
         Width           =   615
      End
      Begin VB.Line Line18 
         X1              =   3600
         X2              =   3480
         Y1              =   2520
         Y2              =   2400
      End
      Begin VB.Line Line17 
         X1              =   3240
         X2              =   3360
         Y1              =   2520
         Y2              =   2400
      End
      Begin VB.Line Line16 
         X1              =   3600
         X2              =   3480
         Y1              =   2280
         Y2              =   2400
      End
      Begin VB.Line Line15 
         X1              =   3480
         X2              =   3720
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line14 
         X1              =   3240
         X2              =   3360
         Y1              =   2280
         Y2              =   2400
      End
      Begin VB.Line Line13 
         X1              =   3120
         X2              =   3360
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label2 
         Caption         =   "Motif 1:"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Motif 2:"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Motif 3:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Window step:"
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Dinucleotide:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sliding window length:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   240
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   949
      TabIndex        =   1
      Top             =   360
      Width           =   14295
      Begin VB.Shape Window_Shape 
         BorderColor     =   &H00808080&
         Height          =   1455
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label Status_sus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Processing: 100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Left            =   5640
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   3840
      End
   End
   Begin VB.TextBox secventata 
      Height          =   1695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2280
      Width           =   11175
   End
   Begin VB.Label Label17 
      Caption         =   "Sequence name :"
      Height          =   255
      Left            =   240
      TabIndex        =   76
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label seq_name_from_file 
      Caption         =   "DNA sequence name ..."
      Height          =   255
      Left            =   1560
      TabIndex        =   75
      Top             =   2040
      Width           =   9975
   End
   Begin VB.Image graf_point_I 
      Height          =   135
      Left            =   12720
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label info 
      Caption         =   "Overall results ..."
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "PromKappa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ________________________________                          _____________________
' /  PromKappa                     \________________________/       v2.00         |
' |                                                                               |
' |            Name:  PromKappa                                                   |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |                                                                               |
' |    Date Created:  3/01/2012                                                   |
' |       Tested On:  WinXP, WinVista, Win7, Win8                                 |
' |             Use:  Analysis of gene promoters                                  |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|

Dim oprire As Boolean
Dim numar As Variant

Dim total_val_x As Variant
Dim total_val_y As Variant

Dim Pause_all_tests As Boolean

Dim pascupas As Variant

Private Sub DubleN_LostFocus()
DINUCLEO.Caption = DubleN.Text
End Sub

Private Sub Erase_Center_patt_Click()
Center_patt.Cls
End Sub

Private Sub Erase_MasterGraph_Click()
MasterGraph.Cls
End Sub

Private Sub Filtru1_Click()
secventata.Text = Replace(secventata.Text, Chr(10), "")
secventata.Text = Replace(secventata.Text, Chr(13), "")
End Sub

Private Sub Filtru2_Click()
secventata.Text = UCase(secventata.Text)
End Sub

Private Sub Filtru3_Click()
secventata.Text = LCase(secventata.Text)
End Sub

Private Sub Filtru4_Click()
Dim ans As String
ans = InputBox("Enter the number of nucleotide (DNA sequence length)", "Random DNA sequence generator", 10000)
If ans = "" Then
MsgBox "Random DNA sequence has not been generated"
Else
secventata.Text = GENEREAZA_NUCLEOTIDE(ans, "ADN")
End If

End Sub

Private Sub Form_Load()
pascupas = 0
  dlgBrowse.InitDir = App.Path 'GetPrimaryDrive
  dlgBrowse.Filter = "(All files)|*.*|"
  dlgBrowse.Flags = cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist + cdlOFNFileMustExist

aq = 0
numar = 0

total_val_x = 0 ' valori totale per set de promotori
total_val_y = 0

Filtru1_Click

DubleN.AddItem "AT"
DubleN.AddItem "AC"
DubleN.AddItem "AG"
DubleN.AddItem "AA"

DubleN.AddItem "CA"
DubleN.AddItem "CT"
DubleN.AddItem "CG"
DubleN.AddItem "CC"

DubleN.AddItem "GT"
DubleN.AddItem "GA"
DubleN.AddItem "GC"
DubleN.AddItem "GG"

DubleN.AddItem "TC"
DubleN.AddItem "TG"
DubleN.AddItem "TA"
DubleN.AddItem "TT"

Pause_all_tests = False
End Sub

Private Sub gene_search_Click()
pro_no = List_promotori.ListCount

For i = 0 To pro_no

    If List_promotori.List(i) <> "" Then

        If InStr(UCase(List_promotori.List(i)), UCase(gene_search_txt.Text)) <> 0 Then
        
        a = List_promotori.ItemData(i)
        MsgBox "Item found in row no " & i
        
            List_promotori.Enabled = False

            secventata.Text = Split(List_promotori.List(i), "|")(1)
            seq_name_from_file.Caption = Split(List_promotori.List(i), "|")(0)
            Procesare_Click

            List_promotori.Enabled = True
        
        
        
        Exit Sub
        End If
        

    End If


Next i


MsgBox "Gene promoter not found !"
End Sub

Private Sub Graph1_show_Click()
MRP.Visible = True
End Sub

Private Sub Graph2_show_Click()
MRP.Visible = False
End Sub

Private Sub Pause_p_test_Click()
If Pause_all_tests = True Then
    Pause_all_tests = False
    Else
    Pause_all_tests = True
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Window_Shape.Visible = True
On Error Resume Next
q = (Len(secventata.Text) / Picture1.ScaleWidth) * X
secventata.SetFocus
secventata.SelStart = q
secventata.SelLength = Lungime_fereastra.Text

secventaADN = secventata.SelText

If Len(secventata.Text) < 3 Then Exit Sub

Window_Shape.Width = (Picture1.ScaleWidth / Len(secventata.Text)) * Len(secventaADN)
Window_Shape.Left = X

If Len(secventaADN) >= Val(Lungime_fereastra.Text) Then
For i = 1 To Val(Lungime_fereastra.Text)
nucleotida = LCase(Mid(secventaADN, i, 1))
If nucleotida = "a" Then a = a + 1
If nucleotida = "t" Then t = t + 1
If nucleotida = "g" Then G = G + 1
If nucleotida = "c" Then c = c + 1
Next i

Toltal_CG_Procent = (100 / (c + G + t + a)) * (c + G)

If SecSenzitiv.Value = 0 Then

Focus_Shape.Width = (graf_point.ScaleWidth / Val(Lungime_fereastra.Text))
Focus_Shape.Height = 3
Focus_Shape.Left = ((graf_point.ScaleWidth / 100) * Toltal_CG_Procent)
Focus_Shape.Top = graf_point.ScaleHeight - ((graf_point.ScaleHeight / 100) * IC(secventaADN)) - 1

Else

Focus_Shape.Width = Xa
Focus_Shape.Height = 3
Focus_Shape.Left = ((graf_point.ScaleWidth / 100) * Toltal_CG_Procent)
Focus_Shape.Top = graf_point.ScaleHeight - ((graf_point.ScaleHeight / 100) * IC(secventaADN)) - 1

End If

If F_point.Value = 1 Then
Line1.X1 = Focus_Shape.Left
Line1.X2 = Line1.X1
Line5.Y1 = Focus_Shape.Top
Line5.Y2 = Line5.Y1
End If

End If

info.Caption = "Sliding window starts at ~ : " & Int(q) & "b and ends at " & _
Int(q + Val(Lungime_fereastra.Text)) & "b" & vbCrLf
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
q = (Len(secventata.Text) / Picture1.ScaleWidth) * X
secventata.SetFocus
secventata.SelStart = q
secventata.SelLength = Lungime_fereastra.Text
End Sub


Function call_sub_Procesare_Click()
Procesare_Click
End Function



Private Sub Procesare_Click()
Dim punct As Long

oprire = False
Window_Shape.Visible = False

mot1.Text = "0"
mot11.Text = "0"
mot2.Text = "0"
mot22.Text = "0"
mot3.Text = "0"
mot33.Text = "0"

If Check1.Value = 1 Then
Picture1.Cls
graf_point.Cls
graf_DI.Cls
graf_IC_CG.Cls
graf_TM.Cls
End If

Status_sus.Visible = True
'-----------------------------------------------------------------
lungimeADN = Len(secventata.Text)

secventaADN = Replace(secventata.Text, vbCrLf, "")
Fereastra = Val(Lungime_fereastra.Text)

For i = 1 To lungimeADN
nucleotida = LCase(Mid(secventaADN, i, 1))
If nucleotida = "a" Then a = a + 1
If nucleotida = "t" Then t = t + 1
If nucleotida = "g" Then G = G + 1
If nucleotida = "c" Then c = c + 1
Next i

Toltal_CG_Procent = (100 / (c + G + t + a)) * (c + G)

Rezultate_txt.Text = Rezultate_txt.Text & "CG = " & Toltal_CG_Procent & " %" & vbCrLf & vbCrLf & _
"A = " & Int((100 / (c + G + t + a)) * a) & " %" & vbCrLf & _
"T = " & Int((100 / (c + G + t + a)) * t) & " %" & vbCrLf & _
"C = " & Int((100 / (c + G + t + a)) * c) & " %" & vbCrLf & _
"G= " & Int((100 / (c + G + t + a)) * G) & " %" & vbCrLf
'-----------------------------------------------------------------

For i = 1 To lungimeADN - Fereastra Step Val(Step_Window.Text)

Motifs_pos = Val(Picture1.ScaleWidth) / Val(lungimeADN)

a = 0
t = 0
c = 0
G = 0

Status_sus.Caption = "Processing: " & Int((100 / (lungimeADN - Fereastra)) * i) & " %"


For j = 1 To Fereastra


nucleotida = LCase(Mid(secventaADN, i + j - 1, 1))

If nucleotida = "a" Then a = a + 1
If nucleotida = "t" Then t = t + 1
If nucleotida = "g" Then G = G + 1
If nucleotida = "c" Then c = c + 1


fereastra_continut = fereastra_continut & nucleotida

Next j

If oprire = True Then Exit Sub
DoEvents

If SecSenzitiv.Value = 1 Then
Fereastra_CG_Procent = (Toltal_CG_Procent / (c + G + t + a)) * (c + G)
Else
Fereastra_CG_Procent = (100 / (c + G + t + a)) * (c + G)
End If



Sir_Proportii = Sir_Proportii & Fereastra_CG_Procent & ","


If PROC1.Value = 1 Then
formula_primer_PCR = Int(81.5 + 16.6 * (Log(0.05) / Log(10)) + 0.41 * (Fereastra_CG_Procent) - 675 / Len(fereastra_continut))
Sir_TM = Sir_TM & formula_primer_PCR & ","
End If

If PROC2.Value = 1 Then
Fereastra_Kappa = IC(fereastra_continut)
Sir_IC = Sir_IC & Fereastra_Kappa & ","
End If

If PROC3.Value = 1 Then
Fereastra_CG_Content = Dinucleotide_Content(fereastra_continut)
Sir_DI = Sir_DI & Fereastra_CG_Content & ","
End If


fereastra_continut_strand2 = strand2(fereastra_continut)


If Mid(UCase(fereastra_continut), 1, Len(TATA.Text)) = TATA.Text Then
mi = Val(Motifs_pos) * i
Picture1.Line (mi, 0)-(mi, 30), vbRed
Picture1.Line (mi, 30)-(mi + 5, 30), vbRed
Picture1.Line (mi + 5, 30)-(mi + 2, 27), vbRed
Picture1.Line (mi + 5, 30)-(mi + 2, 33), vbRed

mot1.Text = Val(mot1.Text) + 1

End If

If Mid(UCase(fereastra_continut_strand2), 1, Len(TATA.Text)) = TATA.Text Then ' STRAND 2
mi = Val(Motifs_pos) * i
Picture1.Line (mi, 0)-(mi, 30), vbRed
Picture1.Line (mi, 30)-(mi - 5, 30), vbRed
Picture1.Line (mi - 5, 30)-(mi - 2, 27), vbRed
Picture1.Line (mi - 5, 30)-(mi - 2, 33), vbRed

mot11.Text = Val(mot11.Text) + 1

End If

If Mid(UCase(fereastra_continut), 1, Len(ATG.Text)) = ATG.Text Then
mi = Val(Motifs_pos) * i
Picture1.Line (mi, 0)-(mi, 20), &HC000&
Picture1.Line (mi, 20)-(mi + 5, 20), &HC000&
Picture1.Line (mi + 5, 20)-(mi + 2, 17), &HC000&
Picture1.Line (mi + 5, 20)-(mi + 2, 23), &HC000&

mot2.Text = Val(mot2.Text) + 1

End If

If Mid(UCase(fereastra_continut_strand2), 1, Len(ATG.Text)) = ATG.Text Then ' STRAND 2
mi = Val(Motifs_pos) * i
Picture1.Line (mi, 0)-(mi, 20), &HC000&
Picture1.Line (mi, 20)-(mi - 5, 20), &HC000&
Picture1.Line (mi - 5, 20)-(mi - 2, 17), &HC000&
Picture1.Line (mi - 5, 20)-(mi - 2, 23), &HC000&

mot22.Text = Val(mot22.Text) + 1

End If

If Mid(UCase(fereastra_continut), 1, Len(ATAAA.Text)) = ATAAA.Text Then
mi = Val(Motifs_pos) * i
Picture1.Line (mi, 0)-(mi, 10), vbBlue
Picture1.Line (mi, 10)-(mi + 5, 10), vbBlue
Picture1.Line (mi + 5, 10)-(mi + 2, 7), vbBlue
Picture1.Line (mi + 5, 10)-(mi + 2, 13), vbBlue

mot3.Text = Val(mot3.Text) + 1

End If

If Mid(UCase(fereastra_continut_strand2), 1, Len(ATAAA.Text)) = ATAAA.Text Then ' STRAND 2
mi = Val(Motifs_pos) * i
Picture1.Line (mi, 0)-(mi, 10), vbBlue
Picture1.Line (mi, 10)-(mi - 5, 10), vbBlue
Picture1.Line (mi - 5, 10)-(mi - 2, 7), vbBlue
Picture1.Line (mi - 5, 10)-(mi - 2, 13), vbBlue

mot33.Text = Val(mot33.Text) + 1

End If

fereastra_continut = ""

Next i

If PROC1.Value = 1 Then Call Deseneaza_TM(lungimeADN, Sir_TM)
If PROC2.Value = 1 Then Call Deseneaza_grafic_IC(lungimeADN, Sir_IC)
If PROC3.Value = 1 Then Call Deseneaza_DI(lungimeADN, Sir_DI)
If PROC4.Value = 1 Then Call Deseneaza_grafic(lungimeADN, Sir_Proportii)


Dim aa() As String ' C+G procentage array
Dim bb() As String ' Kappa IC array
Dim cc() As String ' Dinucleotide content array
Dim tt() As String ' TM array

aa() = Split(Sir_Proportii, ",")
bb() = Split(Sir_IC, ",")
cc() = Split(Sir_DI, ",")
tt() = Split(Sir_TM, ",")

Xa = graf_point.ScaleWidth / 100
Yb = graf_point.ScaleHeight / 100
Xc = graf_DI.ScaleWidth / 100 'dinucleotide

For i = 1 To UBound(aa()) - 1
    
mean_aa = mean_aa + Val(aa(i))
mean_bb = mean_bb + Val(bb(i))
'mean_cc = mean_cc + Val(cc(i)) 'dinucleotide
'mean_tt = mean_tt + Val(tt(i))
    
    If PROC4.Value = 1 And PROC2.Value = 1 Then

    punct = graf_point.Point(aa(i) * Xa, graf_point.ScaleHeight - (bb(i) * Yb))
    
    If punct = Col_pic3.BackColor Or punct = Empty Then punct = Col_pic1.BackColor
    punct = BlendColors(punct, Col_pic2.BackColor, 0.9)
    
If SecSenzitiv.Value = 1 Then
BAR_WITH = Xa
Else
BAR_WITH = (graf_point.ScaleWidth / Val(Lungime_fereastra.Text)) - 1
End If
    
    graf_point.Line (aa(i) * Xa, graf_point.ScaleHeight - (bb(i) * Yb))-((aa(i) * Xa) + (BAR_WITH), graf_point.ScaleHeight - (bb(i) * Yb)), punct
    End If
            
    If PROC4.Value = 1 And PROC2.Value = 1 Then
        If aa(i) > bb(i) Then
            aaa = aaa + 1
            graf_IC_CG.Line (aa(i) * Xa, graf_IC_CG.ScaleHeight - (bb(i) * Yb))-((aa(i) * Xa) + (BAR_WITH), graf_IC_CG.ScaleHeight - (bb(i) * Yb)), BlendColorsRGB(55, 0, 0, 22, 0, 0, vbRed)
        Else
            aaa = aaa - 1
            graf_IC_CG.Line (aa(i) * Xa, graf_IC_CG.ScaleHeight - (bb(i) * Yb))-((aa(i) * Xa) + (BAR_WITH), graf_IC_CG.ScaleHeight - (bb(i) * Yb)), vbBlue
        End If
    End If
            
        If PROC3.Value = 1 And PROC2.Value = 1 Then
        graf_DI.Line (cc(i) * Xc, graf_DI.ScaleHeight - (bb(i) * Yb))-((cc(i) * Xc) + (BAR_WITH), graf_DI.ScaleHeight - (bb(i) * Yb)), vbRed
        End If
        
        If PROC1.Value = 1 And PROC2.Value = 1 Then graf_TM.Line (tt(i) * Xa, graf_TM.ScaleHeight - (bb(i) * Yb))-((tt(i) * Xa) + (BAR_WITH), graf_TM.ScaleHeight - (bb(i) * Yb)), vbRed
       
Next i

mean_aa = mean_aa / UBound(aa())
mean_bb = mean_bb / UBound(bb())

graf_point.Circle (mean_aa * Xa, graf_point.ScaleHeight - (mean_bb * Yb)), 6, vbBlack

'---------------------------------------------------------------------------
total_val_x = total_val_x + Val(mean_aa * Xa)
total_val_y = total_val_y + Val(Center_patt.ScaleHeight - (mean_bb * Yb))
'---------------------------------------------------------------------------
If Option1.Value = True Then
    punct = Center_patt.Point(mean_aa * Xa, Center_patt.ScaleHeight - (mean_bb * Yb))
    If punct = Col_pic3.BackColor Or punct = Empty Then punct = Col_pic1.BackColor
    punct = BlendColors(punct, Col_pic2.BackColor, 0.9)
    Center_patt.PSet (mean_aa * Xa, Center_patt.ScaleHeight - (mean_bb * Yb)), punct
End If

If Option2.Value = True Then
    punct = Center_patt.Point(mean_aa * Xa, Center_patt.ScaleHeight - (mean_bb * Yb))
    If punct = Col_pic3.BackColor Or punct = Empty Then punct = Col_pic1.BackColor
    punct = BlendColors(punct, Col_pic2.BackColor, 0.9)
    Center_patt.Circle (mean_aa * Xa, Center_patt.ScaleHeight - (mean_bb * Yb)), 1, punct
End If

On Error Resume Next
If PROC4.Value = 1 And PROC2.Value = 1 Then
Rezultate_txt.Text = Rezultate_txt.Text & vbCrLf & "Intersections Kappa IC with C+G content (IC_CG):" & aaa & " times " & vbCrLf & "DNA length:" & lungimeADN & "b " & vbCrLf
Rezultate_txt.Text = Rezultate_txt.Text & vbCrLf & "(IC_CG / L_DNA) = " & (lungimeADN / aaa)
End If

Status_sus.Visible = False

If PPDOG.Value = 1 Then
    graf_point.ForeColor = vbRed
    graf_point.FontSize = 10
    graf_point.CurrentX = 30
    graf_point.CurrentY = 10
    graf_point.Print seq_name_from_file.Caption

    graf_IC_CG.ForeColor = vbRed
    graf_IC_CG.FontSize = 10
    graf_IC_CG.CurrentX = 30
    graf_IC_CG.CurrentY = 10
    graf_IC_CG.Print seq_name_from_file.Caption
End If

If Save_Result.Value = 1 Then

gene_name_for_file = Split(seq_name_from_file.Caption, ")")(1)
gene_name_for_file = Split(gene_name_for_file, ";")(0)
gene_name_for_file = Replace(gene_name_for_file, "'", "-")
gene_name_for_file = Replace(gene_name_for_file, ":", "-")

    numar = numar + 1
    graf_point_I.Picture = graf_point.Image
    'tmp_name = App.Path & "\chart\" & numar & "-" & Replace(Time, ":", "-") & ".bmp"
    tmp_name = App.Path & "\chart\" & gene_name_for_file & "-[" & numar & "].bmp"
    SavePicture graf_point_I.Picture, tmp_name
    
    
    numar = numar + 1
    graf_point_I.Picture = graf_IC_CG.Image
    tmp_name = App.Path & "\chart_comp\" & gene_name_for_file & "-[" & numar & "].bmp"
    SavePicture graf_point_I.Picture, tmp_name
    
End If

Window_Shape.Visible = True
End Sub


Function Deseneaza_grafic(ByVal lungime As Variant, ByVal sir As Variant)
Dim a() As String

a() = Split(sir, ",")

X = Picture1.ScaleWidth / UBound(a())
    
        For i = 1 To UBound(a()) - 1

            Picture1.Line (i * X, Picture1.ScaleHeight - a(i - 1))-((i + 1) * X, Picture1.ScaleHeight - a(i)), CG_color_line.BackColor
    
        Next i

End Function

Function Deseneaza_DI(ByVal lungime As Variant, ByVal sir As Variant)
Dim a() As String

a() = Split(sir, ",")

X = Picture1.ScaleWidth / UBound(a())
    
        For i = 1 To UBound(a()) - 1

            Picture1.Line (i * X, Picture1.ScaleHeight - a(i - 1))-((i + 1) * X, Picture1.ScaleHeight - a(i)), DI_color_line.BackColor
    
        Next i

End Function

Function Deseneaza_TM(ByVal lungime As Variant, ByVal sir As Variant)
Dim a() As String

a() = Split(sir, ",")

X = Picture1.ScaleWidth / UBound(a())
    
        For i = 1 To UBound(a()) - 1
    
            Picture1.Line (i * X, Picture1.ScaleHeight - a(i - 1))-((i + 1) * X, Picture1.ScaleHeight - a(i)), TM_color_line.BackColor
    
        Next i
    
End Function

Function Deseneaza_grafic_IC(ByVal lungime As Variant, ByVal sir As Variant)
Dim a() As String

a() = Split(sir, ",")

X = Picture1.ScaleWidth / UBound(a())
    
        For i = 1 To UBound(a()) - 1
    
            Picture1.Line (i * X, Picture1.ScaleHeight - a(i - 1))-((i + 1) * X, Picture1.ScaleHeight - a(i)), Kappa_color_line.BackColor
    
        Next i
    
End Function

Private Sub Search_gene_lis_Click()
Search_gene_names.Show
End Sub

Private Sub Stop_procesare_Click()
oprire = True
End Sub

Function Dinucleotide_Content(ByVal Window As String) As String

    CG_nr = Split(Window, LCase(DubleN.Text))
    CG_nr_buff = UBound(CG_nr)

    op = CG_nr_buff * 2
    Total_CG = (100 / Len(Window)) * op

Dinucleotide_Content = Total_CG
    
End Function

Function GENEREAZA_NUCLEOTIDE(ByVal nr As Variant, ByVal tip As String) As String
'***
Dim nucleo(1 To 5) As String
nucleo(1) = "A"
nucleo(2) = "T"
nucleo(3) = "G"
nucleo(4) = "C"
nucleo(5) = "U"

For n = 1 To nr

If (tip = "ADN") Then
c = Int(3 * Rnd(3))
P = P & nucleo(c + 1)
End If

If (tip = "ARN") Then
c = Int(4 * Rnd(4))
If (c + 1 = 2) Then c = 4
P = P & nucleo(c + 1)
End If

Next n
'***
GENEREAZA_NUCLEOTIDE = P
End Function


Private Sub Promotor_Click()
Dim Msg As String

  On Error Resume Next
  dlgBrowse.ShowOpen
  List_promotori.Clear
  
  Msg = dlgBrowse.FileName

If Msg <> "" Then
  If Err.Number = 0 Then
  
    ff = FreeFile
    Open Msg For Input As #ff
        Do While Not EOF(ff)
           Line Input #ff, inputdata
           
           If UBound(Split(inputdata, ">")) > 0 Then


                If tmp <> "" Then
                
                    If InStr(tmp, "N") <> 0 Then
                        
                    Else
                        List_promotori.AddItem nume_gena & "|" & tmp
                    End If
                
                End If
                
                nume_gena = Replace(inputdata, "|", "-")
                tmp = Empty
           
           Else
           
                 tmp = tmp & UCase(inputdata)
           
           End If
           
        Loop
    Close #ff
    
  End If
  
  Frame4.Caption = "Gene promotors(" & List_promotori.ListCount & ")"
  
End If
End Sub

Private Sub Col_1_Click()
dlgBrowse.ShowColor
Col_pic1.BackColor = dlgBrowse.Color
End Sub

Private Sub Col_2_Click()
dlgBrowse.ShowColor
Col_pic2.BackColor = dlgBrowse.Color
End Sub

Private Sub Col_3_Click()
dlgBrowse.ShowColor
Col_pic3.BackColor = dlgBrowse.Color
graf_point.BackColor = Col_pic3.BackColor
End Sub

Private Sub List_promotori_Click()
List_promotori.Enabled = False
secventata.Text = Split(List_promotori.List(List_promotori.ListIndex), "|")(1)
seq_name_from_file.Caption = Split(List_promotori.List(List_promotori.ListIndex), "|")(0)
Procesare_Click
List_promotori.Enabled = True
End Sub


Private Sub Kappa_for_promotor_list_Click()
Stop_p_test.Enabled = True
List_promotori.Enabled = False


total_val_x = 0
total_val_y = 0

pro_no = List_promotori.ListCount

For i = pascupas To pro_no
pascupas = i
    If Stop_p_test.Enabled = False Then Exit For
'-------------- pause ---------------------------------------------
1:
    DoEvents
    If Pause_all_tests = True Then GoTo 1
'-------------- pause ---------------------------------------------
    
    If List_promotori.List(i) <> "" Then

        secventata.Text = Split(List_promotori.List(i), "|")(1)
        seq_name_from_file.Caption = Split(List_promotori.List(i), "|")(0)

    End If

Procesare_Click

Frame4.Caption = "Gene promotors(" & pro_no & "): " & Int((100 / pro_no) * i) & " %"
DoEvents
Next i


x_set = total_val_x / pro_no
y_set = total_val_y / pro_no

If Show_center_PC.Value = 1 Then Center_patt.Circle (x_set, y_set), 6, vbRed
MasterGraph.Circle (x_set, y_set), 6, vbRed

If Save_Result_MP.Value = 1 Then
    numar = numar + 1
    graf_point_I.Picture = Center_patt.Image
    tmp_name = App.Path & "\chart\" & numar & "-" & Replace(Time, ":", "-") & ".bmp"
    SavePicture graf_point_I.Picture, tmp_name
End If

If Save_Result_MR.Value = 1 Then
    numar = numar + 1
    graf_point_I.Picture = MasterGraph.Image
    tmp_name = App.Path & "\chart\" & numar & "-" & Replace(Time, ":", "-") & ".bmp"
    SavePicture graf_point_I.Picture, tmp_name
End If


List_promotori.Enabled = True
End Sub

Private Sub Stop_p_test_Click()
Stop_p_test.Enabled = False
End Sub

Private Sub TM_color_line_Click()
dlgBrowse.ShowColor
TM_color_line.BackColor = dlgBrowse.Color
End Sub

Private Sub Kappa_color_line_Click()
dlgBrowse.ShowColor
Kappa_color_line.BackColor = dlgBrowse.Color
End Sub

Private Sub DI_color_line_Click()
dlgBrowse.ShowColor
DI_color_line.BackColor = dlgBrowse.Color
End Sub

Private Sub CG_color_line_Click()
dlgBrowse.ShowColor
CG_color_line.BackColor = dlgBrowse.Color
End Sub
