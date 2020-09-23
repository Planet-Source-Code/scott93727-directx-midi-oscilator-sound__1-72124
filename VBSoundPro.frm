VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "VB Sound Pro Ver. 1.6"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   11790
   Icon            =   "VBSoundPro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "Presets"
      Height          =   4095
      Left            =   10200
      TabIndex        =   191
      Top             =   120
      Width           =   1215
      Begin VB.TextBox Text37 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   221
         Top             =   3720
         Width           =   255
      End
      Begin VB.TextBox Text36 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   220
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox Text35 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   219
         Top             =   3240
         Width           =   255
      End
      Begin VB.TextBox Text34 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   218
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox Text33 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   217
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox Text32 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   216
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox Text31 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   215
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox Text30 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   214
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox Text29 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   213
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox Text28 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   212
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox Text27 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   211
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox Text26 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   210
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox Text25 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   209
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text24 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   208
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   720
         TabIndex        =   193
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label34 
         Caption         =   "P15"
         Height          =   255
         Left            =   240
         TabIndex        =   207
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label33 
         Caption         =   "P14"
         Height          =   255
         Left            =   240
         TabIndex        =   206
         Top             =   3480
         Width           =   375
      End
      Begin VB.Label Label32 
         Caption         =   "P13"
         Height          =   255
         Left            =   240
         TabIndex        =   205
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label31 
         Caption         =   "P12"
         Height          =   255
         Left            =   240
         TabIndex        =   204
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label30 
         Caption         =   "P11"
         Height          =   255
         Left            =   240
         TabIndex        =   203
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label29 
         Caption         =   "P10"
         Height          =   255
         Left            =   240
         TabIndex        =   202
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label28 
         Caption         =   "P09"
         Height          =   255
         Left            =   240
         TabIndex        =   201
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "P08"
         Height          =   255
         Left            =   240
         TabIndex        =   200
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label26 
         Caption         =   "P07"
         Height          =   255
         Left            =   240
         TabIndex        =   199
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label25 
         Caption         =   "P06"
         Height          =   255
         Left            =   240
         TabIndex        =   198
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "P05"
         Height          =   255
         Left            =   240
         TabIndex        =   197
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "P04"
         Height          =   255
         Left            =   240
         TabIndex        =   196
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "P03"
         Height          =   255
         Left            =   240
         TabIndex        =   195
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "P02"
         Height          =   255
         Left            =   240
         TabIndex        =   194
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "P01"
         Height          =   255
         Left            =   240
         TabIndex        =   192
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "MiDi Sound Presets"
      Height          =   615
      Left            =   120
      TabIndex        =   175
      Top             =   7080
      Width           =   11295
      Begin VB.CommandButton Command40 
         Caption         =   "P15"
         Height          =   255
         Left            =   10320
         TabIndex        =   190
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command39 
         Caption         =   "P14"
         Height          =   255
         Left            =   9600
         TabIndex        =   189
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command38 
         Caption         =   "P13"
         Height          =   255
         Left            =   8880
         TabIndex        =   188
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command37 
         Caption         =   "P12"
         Height          =   255
         Left            =   8160
         TabIndex        =   187
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command36 
         Caption         =   "P11"
         Height          =   255
         Left            =   7440
         TabIndex        =   186
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command35 
         Caption         =   "P10"
         Height          =   255
         Left            =   6720
         TabIndex        =   185
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command34 
         Caption         =   "P09"
         Height          =   255
         Left            =   6000
         TabIndex        =   184
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command33 
         Caption         =   "P08"
         Height          =   255
         Left            =   5280
         TabIndex        =   183
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command32 
         Caption         =   "P07"
         Height          =   255
         Left            =   4560
         TabIndex        =   182
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command31 
         Caption         =   "P06"
         Height          =   255
         Left            =   3840
         TabIndex        =   181
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command30 
         Caption         =   "P05"
         Height          =   255
         Left            =   3120
         TabIndex        =   180
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command29 
         Caption         =   "P04"
         Height          =   255
         Left            =   2400
         TabIndex        =   179
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command28 
         Caption         =   "P03"
         Height          =   255
         Left            =   1680
         TabIndex        =   178
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command27 
         Caption         =   "P02"
         Height          =   255
         Left            =   960
         TabIndex        =   177
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command26 
         Caption         =   "P01"
         Height          =   255
         Left            =   240
         TabIndex        =   176
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Status Lights"
      Height          =   4095
      Left            =   8640
      TabIndex        =   140
      Top             =   120
      Width           =   1455
      Begin VB.TextBox Text22 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   174
         Top             =   3720
         Width           =   255
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   173
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   172
         Top             =   3240
         Width           =   255
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   171
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   170
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   169
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   158
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   156
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   155
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   152
         Top             =   1560
         Width           =   255
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   151
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   148
         Top             =   1080
         Width           =   255
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   147
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   144
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FF0000&
         Height          =   195
         Left            =   1080
         TabIndex        =   142
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "MCIClose"
         Height          =   255
         Left            =   120
         TabIndex        =   168
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "MCIOpen"
         Height          =   255
         Left            =   120
         TabIndex        =   167
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "SWCall"
         Height          =   255
         Left            =   120
         TabIndex        =   166
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Song Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Song Play"
         Height          =   255
         Left            =   120
         TabIndex        =   164
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Key Up"
         Height          =   255
         Left            =   120
         TabIndex        =   163
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Key down"
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "MCI Record"
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "MCI Play"
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "PlayBack"
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Record"
         Height          =   255
         Left            =   120
         TabIndex        =   149
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "MiDi Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   146
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "MiDi Play"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Gen Off"
         Height          =   255
         Left            =   120
         TabIndex        =   143
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Gen On"
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Timer tmrplayback 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11400
      Top             =   6840
   End
   Begin VB.Frame Frame8 
      Caption         =   "KeyBoard Keys"
      Height          =   1335
      Left            =   120
      TabIndex        =   68
      Top             =   5640
      Width           =   11295
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   70
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   68
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   66
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   71
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   139
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   69
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   137
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   63
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   131
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   61
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   58
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   56
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   124
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   54
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   51
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   49
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   46
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   44
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   42
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   39
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   32
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   30
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   67
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   135
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   37
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   34
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   25
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   27
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   22
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   20
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   18
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   15
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   13
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   10
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00000000&
         Height          =   495
         Index           =   8
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   6
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   3
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00404040&
         Height          =   495
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   65
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   133
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   64
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   132
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   62
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   130
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   60
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   128
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   59
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   127
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   57
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   125
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   55
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   123
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   53
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   121
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   52
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   120
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   50
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   118
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   48
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   116
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   47
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   115
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   45
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   113
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   43
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   111
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   41
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   109
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   40
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   108
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   38
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   106
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   36
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   104
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   35
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   103
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   33
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   101
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   31
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   99
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   29
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   97
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   28
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   96
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   26
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   94
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   24
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   92
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   23
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   91
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   21
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   89
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   19
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   87
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   17
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   85
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   16
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   84
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   14
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   82
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   12
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   159
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   11
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   80
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   9
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   78
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   7
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   76
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   5
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   74
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   4
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   73
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   71
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton pkey 
         BackColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   69
         Tag             =   "1"
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "MIDI (Note) Recorder"
      Height          =   735
      Left            =   4800
      TabIndex        =   63
      Top             =   3480
      Width           =   3735
      Begin VB.CommandButton cmdplay 
         BackColor       =   &H0080FF80&
         Caption         =   "Play"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdrec 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Record"
         Height          =   375
         Left            =   1920
         MaskColor       =   &H8000000A&
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Save"
         Height          =   375
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdload 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Load"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   240
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Caption         =   "MCI Media Control"
      Height          =   1215
      Left            =   120
      TabIndex        =   43
      Top             =   4320
      Width           =   11295
      Begin VB.CommandButton Command25 
         Caption         =   "MMov"
         Height          =   375
         Left            =   10440
         TabIndex        =   62
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command24 
         Caption         =   "DigVid"
         Height          =   375
         Left            =   10440
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Vdsc"
         Height          =   375
         Left            =   9720
         TabIndex        =   60
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command22 
         Caption         =   "MIDI"
         Height          =   375
         Left            =   9720
         TabIndex        =   59
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Scan"
         Height          =   375
         Left            =   9000
         TabIndex        =   58
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Wave"
         Height          =   375
         Left            =   9000
         TabIndex        =   57
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command19 
         Caption         =   "VCR"
         Height          =   375
         Left            =   8280
         TabIndex        =   56
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command18 
         Caption         =   "CD"
         Height          =   375
         Left            =   8280
         TabIndex        =   55
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Dat"
         Height          =   375
         Left            =   7560
         TabIndex        =   54
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command16 
         Caption         =   "AVI"
         Height          =   375
         Left            =   7560
         TabIndex        =   53
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Save"
         Height          =   375
         Left            =   3960
         TabIndex        =   52
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Seek"
         Height          =   375
         Left            =   5640
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Back"
         Height          =   375
         Left            =   6480
         TabIndex        =   50
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Forward"
         Height          =   375
         Left            =   6480
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Eject"
         Height          =   375
         Left            =   5640
         TabIndex        =   48
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Open"
         Height          =   375
         Left            =   3960
         TabIndex        =   47
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Play "
         Height          =   375
         Left            =   4800
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Record"
         Height          =   375
         Left            =   4800
         TabIndex        =   45
         Top             =   720
         Width           =   735
      End
      Begin MCI.MMControl MMControl1 
         Height          =   855
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   1508
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
   End
   Begin VB.Timer tmrwelcome 
      Interval        =   100
      Left            =   11400
      Top             =   6360
   End
   Begin VB.Timer tmrrec 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   11400
      Top             =   7320
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frequency Adjust"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   4575
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   25
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Mixer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   4095
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         LargeChange     =   20
         Left            =   240
         Max             =   10000
         Min             =   1
         TabIndex        =   22
         Top             =   360
         Value           =   1
         Width           =   4095
      End
   End
   Begin VB.TextBox txtHundreds 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   18
      Top             =   1320
      Width           =   250
   End
   Begin VB.VScrollBar scrHundreds 
      Height          =   765
      Left            =   600
      Max             =   0
      Min             =   9
      TabIndex        =   20
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtThousands 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   19
      Top             =   1320
      Width           =   250
   End
   Begin VB.TextBox txtTens 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   17
      Top             =   1320
      Width           =   250
   End
   Begin VB.TextBox txtUnits 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   16
      Top             =   1320
      Width           =   250
   End
   Begin VB.TextBox txtDecimal 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   1320
      Width           =   265
   End
   Begin VB.VScrollBar scrThousands 
      Height          =   765
      Left            =   360
      Max             =   0
      Min             =   9
      TabIndex        =   14
      Top             =   1080
      Width           =   255
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   30
      Left            =   600
      TabIndex        =   13
      Top             =   1200
      Width           =   255
   End
   Begin VB.VScrollBar scrTens 
      Height          =   765
      Left            =   840
      Max             =   0
      Min             =   9
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.VScrollBar scrUnits 
      Height          =   765
      Left            =   1080
      Max             =   0
      Min             =   9
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.VScrollBar scrDecimal 
      Height          =   765
      Left            =   1320
      Max             =   0
      Min             =   9
      TabIndex        =   10
      Top             =   1080
      Width           =   265
   End
   Begin VB.Frame Frame3 
      Caption         =   "Function"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   2775
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wave Generator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.OptionButton optNoise 
         Caption         =   "Noise"
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
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optSaw 
         Caption         =   "Saw"
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
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optTriangle 
         Caption         =   "Triangle"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optSquare 
         Caption         =   "Square"
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
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optSine 
         Caption         =   "Sine"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frequency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame Frame5 
      Caption         =   "Midi Note Generator / Piano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   4800
      TabIndex        =   26
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command44 
         Caption         =   "Ent"
         Height          =   255
         Left            =   3000
         TabIndex        =   225
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Ent"
         Height          =   255
         Left            =   3000
         TabIndex        =   224
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Ent"
         Height          =   255
         Left            =   3000
         TabIndex        =   223
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Ent"
         Height          =   255
         Left            =   3000
         TabIndex        =   222
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   161
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar sldChan 
         Height          =   255
         Left            =   240
         Max             =   16
         Min             =   1
         TabIndex        =   160
         Top             =   1920
         Value           =   1
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   38
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Mixer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   37
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stop Song"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   36
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Play Song"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   35
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Play"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.HScrollBar sldVol 
         Height          =   255
         Left            =   240
         Max             =   127
         TabIndex        =   29
         Top             =   1440
         Width           =   1455
      End
      Begin VB.HScrollBar sldInst 
         Height          =   255
         Left            =   240
         Max             =   127
         TabIndex        =   28
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar sldPitch 
         Height          =   255
         Left            =   240
         Max             =   100
         TabIndex        =   27
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "      Channel           Channel"
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
         Left            =   240
         TabIndex        =   162
         Top             =   1720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "         Note                Note"
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
         Left            =   240
         TabIndex        =   41
         Top             =   760
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "       Volume             Volume   "
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
         Left            =   240
         TabIndex        =   40
         Top             =   1240
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "    Instruments      Instruments"
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
         Left            =   240
         TabIndex        =   39
         Top             =   280
         Width           =   2655
      End
   End
   Begin VB.TextBox Text6 
      Height          =   3135
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Text            =   "VBSoundPro.frx":030A
      Top             =   240
      Width           =   3735
   End
   Begin VB.Menu op0 
      Caption         =   "Options"
      Begin VB.Menu op1 
         Caption         =   "DXOsc On"
      End
      Begin VB.Menu op2 
         Caption         =   "DXOsc Off"
      End
      Begin VB.Menu op3 
         Caption         =   "Midi In"
      End
      Begin VB.Menu op4 
         Caption         =   "Midi Out"
      End
      Begin VB.Menu op5 
         Caption         =   "MCI On"
      End
      Begin VB.Menu op6 
         Caption         =   "MCI Off"
      End
      Begin VB.Menu opt7 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu ps0 
      Caption         =   "Presets"
      Begin VB.Menu ps1 
         Caption         =   "Set P01"
      End
      Begin VB.Menu ps2 
         Caption         =   "Set P02"
      End
      Begin VB.Menu ps3 
         Caption         =   "Set P03"
      End
      Begin VB.Menu ps4 
         Caption         =   "Set P04"
      End
      Begin VB.Menu ps5 
         Caption         =   "Set P05"
      End
      Begin VB.Menu ps6 
         Caption         =   "Set P06"
      End
      Begin VB.Menu ps7 
         Caption         =   "Set P07"
      End
      Begin VB.Menu ps8 
         Caption         =   "Set P08"
      End
      Begin VB.Menu ps9 
         Caption         =   "Set P09"
      End
      Begin VB.Menu ps10 
         Caption         =   "Set P10"
      End
      Begin VB.Menu ps11 
         Caption         =   "Set P11"
      End
      Begin VB.Menu ps12 
         Caption         =   "Set P12"
      End
      Begin VB.Menu ps13 
         Caption         =   "Set P13"
      End
      Begin VB.Menu ps14 
         Caption         =   "Set P14"
      End
      Begin VB.Menu ps15 
         Caption         =   "Set P15"
      End
   End
   Begin VB.Menu sw0 
      Caption         =   "Software"
      Begin VB.Menu sw1 
         Caption         =   "SWCALL1"
      End
      Begin VB.Menu sw2 
         Caption         =   "SWCALL2"
      End
      Begin VB.Menu sw3 
         Caption         =   "SWCALL3"
      End
      Begin VB.Menu sw4 
         Caption         =   "SWCALL4"
      End
      Begin VB.Menu sw5 
         Caption         =   "SWRESET"
      End
   End
   Begin VB.Menu ab0 
      Caption         =   "About"
      Begin VB.Menu ab1 
         Caption         =   "Specs"
      End
      Begin VB.Menu ab5 
         Caption         =   "Presets"
      End
      Begin VB.Menu ab2 
         Caption         =   "Instruments"
      End
      Begin VB.Menu ab6 
         Caption         =   "Usage"
      End
      Begin VB.Menu ab4 
         Caption         =   "Contact"
      End
      Begin VB.Menu ab3 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Complete VB Sound Demo, DirectX, MiDi recorder Including Duration, 126 Instruments, 5 Wave Format Generator, MCI Control of AVIVideo, CDAudio, DAT, DigitalVideo, MMMovie, Sequencer, VCR, Videodisc, or WaveAudio
'Author Contact scott93727@aol.com

'If you plan on using this code in your application
'registration is required, for Registration Visit :
'www.sharewaresoftware.info/vbsoundpro/vbsoundpro.htm

' some e of the code was from the below authors
' P I A N O  by Armin Niki
' Update - Apr 20 2006 by Paul Bahlawan



Option Explicit

Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

'Create the DirectSound8 Object

Dim dx As New DirectX8
Dim ds As DirectSound8
Dim dsBuffer As DirectSoundSecondaryBuffer8

Private hmidi As Long
Private baseNote As Long
Private channel As Long
Private volume As Long
Private lNote As Long
Private Playin() As String
Private playinc As Long
Private timers As Long
Private rec As String
Dim midimsg As Long
Dim notep As Long

Dim stur, oncol, offcol
'Declare the variables that will be used

Dim frequency As Double
Dim increment As Double
Dim fileName As String
Dim fileSize As Double
Dim sample As Long
Dim period As Double
Dim state As Integer
Dim bufferptr As Long
Dim inputValue As Double

Dim notetm, onoff, noteimp, asup, asdn, inst
Dim cs1a, cs2a, cs3a, cs4a
Dim cs1b, cs2b, cs3b, cs4b
Dim cs1c, cs2c, cs3c, cs4c
Dim cs1d, cs2d, cs3d, cs4d
Dim cs1e, cs2e, cs3e, cs4e
Dim cs1f, cs2f, cs3f, cs4f
Dim cs1g, cs2g, cs3g, cs4g
Dim cs1h, cs2h, cs3h, cs4h
Dim cs1i, cs2i, cs3i, cs4i
Dim cs1j, cs2j, cs3j, cs4j
Dim cs1k, cs2k, cs3k, cs4k
Dim cs1l, cs2l, cs3l, cs4l
Dim cs1m, cs2m, cs3m, cs4m
Dim cs1n, cs2n, cs3n, cs4n
Dim cs1o, cs2o, cs3o, cs4o
Dim cs1p, cs2p, cs3p, cs4p
Dim cs1q, cs2q, cs3q, cs4q
Dim cs1r, cs2r, cs3r, cs4r
Const Pi = 3.141592654
Const sampleRate = 44100
Const amplitude = 127
Private Sub domusickey(key As Integer)
Dim note As Long
    If key = vbKey1 Then note = 1
    If key = vbKey2 Then note = 3
    If key = vbKey3 Then note = 5
    If key = vbKey4 Then note = 6
    If key = vbKey5 Then note = 8
    If key = vbKey6 Then note = 10
    If key = vbKey7 Then note = 12
    If key = vbKey8 Then note = 13
    If key = vbKey9 Then note = 15
    If key = vbKey0 Then note = 17
    If key = 189 Then note = 18
    If key = 187 Then note = 20
    If key = vbKeyQ Then note = 22
    If key = vbKeyW Then note = 24
    If key = vbKeyE Then note = 25
    If key = vbKeyR Then note = 27
    If key = vbKeyT Then note = 29
    If key = vbKeyY Then note = 30
    If key = vbKeyU Then note = 32
    If key = vbKeyI Then note = 34
    If key = vbKeyO Then note = 36
    If key = vbKeyP Then note = 37
    If key = 219 Then note = 39
    If key = vbKeyA Then note = 41
    If key = vbKeyS Then note = 42
    If key = vbKeyD Then note = 44
    If key = vbKeyF Then note = 46
    If key = vbKeyG Then note = 48
    If key = vbKeyH Then note = 49
    If key = vbKeyJ Then note = 51
    If key = vbKeyK Then note = 53
    If key = vbKeyL Then note = 54
    If key = 186 Then note = 56
    If key = vbKeyZ Then note = 58
    If key = vbKeyX Then note = 60
    If key = vbKeyC Then note = 61
    If key = vbKeyV Then note = 63
    If key = vbKeyB Then note = 65
    If key = vbKeyN Then note = 66
    If key = vbKeyM Then note = 68
    If key = 188 Then note = 70
    If key = 190 Then note = 72
    If Not note = lNote And note Then
        domusic note
    End If
End Sub
'Convert a key to a note, then stop that note
Private Sub domusickeystop(key As Integer)
Dim note As Long
    If key = vbKey1 Then note = 1
    If key = vbKey2 Then note = 3
    If key = vbKey3 Then note = 5
    If key = vbKey4 Then note = 6
    If key = vbKey5 Then note = 8
    If key = vbKey6 Then note = 10
    If key = vbKey7 Then note = 12
    If key = vbKey8 Then note = 13
    If key = vbKey9 Then note = 15
    If key = vbKey0 Then note = 17
    If key = 189 Then note = 18
    If key = 187 Then note = 20
    If key = vbKeyQ Then note = 22
    If key = vbKeyW Then note = 24
    If key = vbKeyE Then note = 25
    If key = vbKeyR Then note = 27
    If key = vbKeyT Then note = 29
    If key = vbKeyY Then note = 30
    If key = vbKeyU Then note = 32
    If key = vbKeyI Then note = 34
    If key = vbKeyO Then note = 36
    If key = vbKeyP Then note = 37
    If key = 219 Then note = 39
    If key = vbKeyA Then note = 41
    If key = vbKeyS Then note = 42
    If key = vbKeyD Then note = 44
    If key = vbKeyF Then note = 46
    If key = vbKeyG Then note = 48
    If key = vbKeyH Then note = 49
    If key = vbKeyJ Then note = 51
    If key = vbKeyK Then note = 53
    If key = vbKeyL Then note = 54
    If key = 186 Then note = 56
    If key = vbKeyZ Then note = 58
    If key = vbKeyX Then note = 60
    If key = vbKeyC Then note = 61
    If key = vbKeyV Then note = 63
    If key = vbKeyB Then note = 65
    If key = vbKeyN Then note = 66
    If key = vbKeyM Then note = 68
    If key = 188 Then note = 70
    If key = 190 Then note = 72
    domusicstop note
End Sub
Private Sub ab1_Click()
Dim Msg, Title
Title = "VB Sound Pro Ver. 1.6 "
Msg = "Complete VB Sound Demo, DirectX, MIDI (127 Instruments), Oscilators, MIDI Recorder Including Duration, 127 Instruments, 5 Wave Format Oscilator, MCI Control of AVIVideo, CDAudio, DAT, DigitalVideo, MMMovie, Sequencer, VCR, Videodisc, or WaveAudio.Best Run at 800X600 Resolution."
MsgBox Msg, , Title
End Sub
Private Sub ab5_Click()
Dim Msg, Title
Title = "VB Sound Pro Ver. 1.6"
Msg = "Presets allow the user to store 15 presets all each with seperate Instrument, Pitch, Volume and Channels Settings, to use adjust the 4 settings (Instrument, Pitch, Volume, Channel), and click the <Set Preset # of your choice> Menu Option under the <Presets> section."
MsgBox Msg, , Title
End Sub
Private Sub ab6_Click()
Dim Msg, Title
Title = "VB Sound Pro Ver. 1.6 "
Msg = "If you plan on using this code in your application, registration is required, Visit www.sharewaresoftware.info/vbsoundpro/vbsoundpro.htm for Registration."
MsgBox Msg, , Title
End Sub
Private Sub Command26_Click()

Text23.BackColor = "255"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"

On Error Resume Next
Open "cust1" For Input As #1
Input #1, cs1a, cs2a, cs3a, cs4a
Close #1
sldVol.Value = cs1a
sldPitch.Value = cs2a
sldChan.Value = cs3a
sldInst.Value = cs4a
End Sub
Private Sub Command27_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "255"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"

On Error Resume Next
Open "cust2" For Input As #1
Input #1, cs1b, cs2b, cs3b, cs4b
Close #1
sldVol.Value = cs1b
sldPitch.Value = cs2b
sldChan.Value = cs3b
sldInst.Value = cs4b
End Sub
Private Sub Command28_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "255"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust3" For Input As #1
Input #1, cs1c, cs2c, cs3c, cs4c
Close #1
sldVol.Value = cs1c
sldPitch.Value = cs2c
sldChan.Value = cs3c
sldInst.Value = cs4c
End Sub
Private Sub Command29_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "255"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust4" For Input As #1
Input #1, cs1d, cs2d, cs3d, cs4d
Close #1
sldVol.Value = cs1d
sldPitch.Value = cs2d
sldChan.Value = cs3d
sldInst.Value = cs4d
End Sub
Private Sub Command30_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "255"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust5" For Input As #1
Input #1, cs1e, cs2e, cs3e, cs4e
Close #1
sldVol.Value = cs1e
sldPitch.Value = cs2e
sldChan.Value = cs3e
sldInst.Value = cs4e
End Sub
Private Sub Command31_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "255"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust6" For Input As #1
Input #1, cs1f, cs2f, cs3f, cs4f
Close #1
sldVol.Value = cs1f
sldPitch.Value = cs2f
sldChan.Value = cs3f
sldInst.Value = cs4f
End Sub
Private Sub Command32_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "255"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust7" For Input As #1
Input #1, cs1g, cs2g, cs3g, cs4g
Close #1
sldVol.Value = cs1g
sldPitch.Value = cs2g
sldChan.Value = cs3g
sldInst.Value = cs4g
End Sub
Private Sub Command33_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "255"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust8" For Input As #1
Input #1, cs1h, cs2h, cs3h, cs4h
Close #1
sldVol.Value = cs1h
sldPitch.Value = cs2h
sldChan.Value = cs3h
sldInst.Value = cs4h
End Sub
Private Sub Command34_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "255"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust9" For Input As #1
Input #1, cs1i, cs2i, cs3i, cs4i
Close #1
sldVol.Value = cs1i
sldPitch.Value = cs2i
sldChan.Value = cs3i
sldInst.Value = cs4i
End Sub
Private Sub Command35_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "255"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust10" For Input As #1
Input #1, cs1j, cs2j, cs3j, cs4j
Close #1
sldVol.Value = cs1j
sldPitch.Value = cs2j
sldChan.Value = cs3j
sldInst.Value = cs4j
End Sub
Private Sub Command36_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "255"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust11" For Input As #1
Input #1, cs1k, cs2k, cs3k, cs4k
Close #1
sldVol.Value = cs1k
sldPitch.Value = cs2k
sldChan.Value = cs3k
sldInst.Value = cs4k
End Sub
Private Sub Command37_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "255"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust12" For Input As #1
Input #1, cs1l, cs2l, cs3l, cs4l
Close #1
sldVol.Value = cs1l
sldPitch.Value = cs2l
sldChan.Value = cs3l
sldInst.Value = cs4l
End Sub
Private Sub Command38_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "255"
Text36.BackColor = "16711680"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust13" For Input As #1
Input #1, cs1m, cs2m, cs3m, cs4m
Close #1
sldVol.Value = cs1m
sldPitch.Value = cs2m
sldChan.Value = cs3m
sldInst.Value = cs4m
End Sub
Private Sub Command39_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "255"
Text37.BackColor = "16711680"
On Error Resume Next
Open "cust14" For Input As #1
Input #1, cs1n, cs2n, cs3n, cs4n
Close #1
sldVol.Value = cs1n
sldPitch.Value = cs2n
sldChan.Value = cs3n
sldInst.Value = cs4n
End Sub
Private Sub Command40_Click()
Text23.BackColor = "16711680"
Text24.BackColor = "16711680"
Text25.BackColor = "16711680"
Text26.BackColor = "16711680"
Text27.BackColor = "16711680"
Text28.BackColor = "16711680"
Text29.BackColor = "16711680"
Text30.BackColor = "16711680"
Text31.BackColor = "16711680"
Text32.BackColor = "16711680"
Text33.BackColor = "16711680"
Text34.BackColor = "16711680"
Text35.BackColor = "16711680"
Text36.BackColor = "16711680"
Text37.BackColor = "255"
On Error Resume Next
Open "cust15" For Input As #1
Input #1, cs1o, cs2o, cs3o, cs4o
Close #1
sldVol.Value = cs1o
sldPitch.Value = cs2o
sldChan.Value = cs3o
sldInst.Value = cs4o
End Sub
Private Sub Command41_Click()
On Error Resume Next
sldInst.Value = Text3.Text
End Sub
Private Sub Command42_Click()
On Error Resume Next
sldPitch.Value = Text4.Text
End Sub
Private Sub Command43_Click()
On Error Resume Next
sldVol.Value = Text5.Text
End Sub
Private Sub Command44_Click()
On Error Resume Next
sldChan.Value = Text16.Text
End Sub
'Play piano via keyboard
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    domusickey KeyCode
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    domusickeystop KeyCode
End Sub
Private Sub MMControl1_Done(NotifyCode As Integer)
Text13.BackColor = "16711680"
Text14.BackColor = "16711680"
End Sub
Private Sub opt7_Click()
On Error Resume Next
    dsBuffer.Stop
    Set dsBuffer = Nothing
    
    On Error Resume Next
    tmrwelcome.Enabled = False
domusicstop notep

    midiOutClose (hmidi)
    MMControl1.Command = "Close"

End Sub
Private Sub sldChan_Change()
    channel = sldChan.Value - 1
    Text16.Text = channel
End Sub
'Play piano via mouse
Private Sub pKey_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text15.BackColor = "255"
    Text17.BackColor = "16711680"
    domusic Index + 1
    
    Text4.Text = Index
End Sub
Private Sub swcall(Index As Integer)
    ' used to call a note via software
    Text20.BackColor = "255"
    domusic Index + 1
    domusicstop Index + 1
    'example below
    'Private Sub Command41_Click()
    'Text20.BackColor = "255"
    'Call swcall(19)
    'End Sub
    'or: if x=5 then call swcall(19)(19 is the note)
End Sub
Private Sub pKey_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text15.BackColor = "16711680"
    Text17.BackColor = "255"
    domusicstop Index + 1
End Sub
Private Sub domusic(mNote As Long)
Dim midimsg As Long
    'Play note
    midimsg = &H90 + ((baseNote + mNote) * &H100) + (volume * &H10000) + channel
    midiOutShortMsg hmidi, midimsg
    'record the key-down event
    If tmrrec.Enabled Then rec = rec & mNote & "x" & timers & " "
    timers = 0
    lNote = mNote
    'hi-light key being played
    On Error Resume Next
    pkey(mNote - 1).BackColor = &HC000&    '&H6060F0
End Sub
'Stop a note
Private Sub domusicstop(mNote As Long)
Dim midimsg As Long
    midimsg = &H80 + ((baseNote + mNote) * &H100) + channel
    midiOutShortMsg hmidi, midimsg
    'record the key-up event
    If tmrrec.Enabled Then rec = rec & -mNote & "x" & timers & " "
    timers = 0
    If mNote = lNote Then lNote = 0 'lNote = 0
    'mNote = 0
    If mNote > 0 Then 'un-hi-light key when released
        If pkey(mNote - 1).Tag = "1" Then
            pkey(mNote - 1).BackColor = vbWhite
        Else
            pkey(mNote - 1).BackColor = vbBlack
        End If
    End If
End Sub
'Show middle C when pitch is changed
Private Sub sldPitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then
        pkey(59 - sldPitch.Value).SetFocus
    End If
End Sub
Private Sub ab2_Click()
If stur = 2 Then stur = 1: GoTo stra:
If stur = 1 Then stur = 2: GoTo strb:

stra:
Text6.Visible = True
Frame5.Visible = False
GoTo stred

strb:
Text6.Visible = False
Frame5.Visible = True
GoTo stred

stred:
End Sub
Private Sub ab3_Click()
Dim Msg, Title
Title = "VB Sound Pro Ver. 1.6"
Msg = "Easy to Understand VB Sound Support for all Sound functions Including DirectX, MIDI, Oscilators, and MCI Functions.(Waves, MP3, MCI, MiDi, DirectX, AVIVideo, CDAudio, DAT, DigitalVideo, MMMovie, Sequencer, VCR, Videodisc, or WaveAudio)."
MsgBox Msg, , Title
End Sub
Private Sub ab4_Click()
Dim Msg, Title
Title = "VB Sound Pro Ver. 1.6 "
Msg = "(c) 2009 PWS software, Tech support scott93727@aol.com, http://www.freeclimate.com/pwssw/pwssw.htm"
MsgBox Msg, , Title
End Sub
Private Sub cmdrecord_Click()
If tmrplayback.Enabled Then cmdPlay_Click
    If tmrrec.Enabled Then
        'stop recording
        tmrrec.Enabled = False
        cmdrec.BackColor = &HC0C0FF
    Else
        'start recording
        rec = ""
        timers = 0
        tmrrec.Enabled = True
        cmdrec.BackColor = &H4040FF
    End If
End Sub
Private Sub Command10_Click()
CommonDialog1.DialogTitle = "Media Player Open"
On Error Resume Next
 CommonDialog1.CancelError = True
    On Error GoTo ErrHandler:
    CommonDialog1.flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
   MMControl1.fileName = CommonDialog1.fileName
    'Me.Caption = "NoteEdit " & CommonDialog1.filename
ErrHandler:
    'if Cancel clicked, then exit procedure
'MMControl1.DeviceType = "WaveAudio"

' Open the MCI WaveAudio device.
   MMControl1.Command = "Open"
End Sub
Private Sub ps1_Click()
cs1a = sldVol.Value
cs2a = sldPitch.Value
cs3a = sldChan.Value
cs4a = sldInst.Value
On Error Resume Next
Open "cust1" For Output As #1
Write #1, cs1a, cs2a, cs3a, cs4a
Close #1
End Sub
Private Sub ps2_Click()
cs1b = sldVol.Value
cs2b = sldPitch.Value
cs3b = sldChan.Value
cs4b = sldInst.Value
On Error Resume Next
Open "cust2" For Output As #1
Write #1, cs1b, cs2b, cs3b, cs4b
Close #1
End Sub
Private Sub ps3_Click()
cs1c = sldVol.Value
cs2c = sldPitch.Value
cs3c = sldChan.Value
cs4c = sldInst.Value
On Error Resume Next
Open "cust3" For Output As #1
Write #1, cs1c, cs2c, cs3c, cs4c
Close #1
End Sub
Private Sub ps4_Click()
cs1d = sldVol.Value
cs2d = sldPitch.Value
cs3d = sldChan.Value
cs4d = sldInst.Value
On Error Resume Next
Open "cust4" For Output As #1
Write #1, cs1d, cs2d, cs3d, cs4d
Close #1
End Sub
Private Sub ps5_Click()
cs1e = sldVol.Value
cs2e = sldPitch.Value
cs3e = sldChan.Value
cs4e = sldInst.Value
On Error Resume Next
Open "cust5" For Output As #1
Write #1, cs1e, cs2e, cs3e, cs4e
Close #1
End Sub
Private Sub ps6_Click()
cs1f = sldVol.Value
cs2f = sldPitch.Value
cs3f = sldChan.Value
cs4f = sldInst.Value
On Error Resume Next
Open "cust6" For Output As #1
Write #1, cs1f, cs2f, cs3f, cs4f
Close #1
End Sub
Private Sub ps7_Click()
cs1g = sldVol.Value
cs2g = sldPitch.Value
cs3g = sldChan.Value
cs4g = sldInst.Value
On Error Resume Next
Open "cust7" For Output As #1
Write #1, cs1g, cs2g, cs3g, cs4g
Close #1
End Sub
Private Sub ps8_Click()
cs1h = sldVol.Value
cs2h = sldPitch.Value
cs3h = sldChan.Value
cs4h = sldInst.Value
On Error Resume Next
Open "cust8" For Output As #1
Write #1, cs1h, cs2h, cs3h, cs4h
Close #1
End Sub
Private Sub ps9_Click()
cs1i = sldVol.Value
cs2i = sldPitch.Value
cs3i = sldChan.Value
cs4i = sldInst.Value
On Error Resume Next
Open "cust9" For Output As #1
Write #1, cs1i, cs2i, cs3i, cs4i
Close #1
End Sub
Private Sub ps10_Click()
cs1j = sldVol.Value
cs2j = sldPitch.Value
cs3j = sldChan.Value
cs4j = sldInst.Value
On Error Resume Next
Open "cust10" For Output As #1
Write #1, cs1j, cs2j, cs3j, cs4j
Close #1
End Sub
Private Sub ps11_Click()
cs1k = sldVol.Value
cs2k = sldPitch.Value
cs3k = sldChan.Value
cs4k = sldInst.Value
On Error Resume Next
Open "cust11" For Output As #1
Write #1, cs1k, cs2k, cs3k, cs4k
Close #1
End Sub
Private Sub ps12_Click()
cs1l = sldVol.Value
cs2l = sldPitch.Value
cs3l = sldChan.Value
cs4l = sldInst.Value
On Error Resume Next
Open "cust12" For Output As #1
Write #1, cs1l, cs2l, cs3l, cs4l
Close #1
End Sub
Private Sub ps13_Click()
cs1m = sldVol.Value
cs2m = sldPitch.Value
cs3m = sldChan.Value
cs4m = sldInst.Value
On Error Resume Next
Open "cust13" For Output As #1
Write #1, cs1m, cs2m, cs3m, cs4m
Close #1
End Sub
Private Sub ps14_Click()
cs1n = sldVol.Value
cs2n = sldPitch.Value
cs3n = sldChan.Value
cs4n = sldInst.Value
On Error Resume Next
Open "cust14" For Output As #1
Write #1, cs1n, cs2n, cs3n, cs4n
Close #1
End Sub
Private Sub ps15_Click()
cs1o = sldVol.Value
cs2o = sldPitch.Value
cs3o = sldChan.Value
cs4o = sldInst.Value
On Error Resume Next
Open "cust15" For Output As #1
Write #1, cs1o, cs2o, cs3o, cs4o
Close #1
End Sub
Private Sub Command11_Click()
On Error Resume Next
MMControl1.Command = "Eject"
End Sub
Private Sub Command12_Click()
On Error Resume Next
MMControl1.Command = "Step"
End Sub
Private Sub Command13_Click()
On Error Resume Next
MMControl1.Command = "Back"
End Sub
Private Sub Command14_Click()
On Error Resume Next
MMControl1.Command = "Seek"
End Sub
Private Sub Command15_Click()
On Error Resume Next
MMControl1.Command = "Save"
End Sub
Private Sub Command16_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "AVIVideo"
End Sub
Private Sub Command17_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "Dat"
End Sub
Private Sub Command18_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "CDAudio"
End Sub
Private Sub Command19_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "VCR"
End Sub
Private Sub Command2_Click()
Dim retval
On Error Resume Next
retval = Shell("c:\windows\sndvol32", 1)
On Error Resume Next
retval = Shell("c:\windows\system32\sndvol32", 1)
End Sub
Private Sub Command20_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "WaveAudio"
End Sub
Private Sub Command21_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "Scanner"
End Sub
Private Sub Command22_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "Sequencer"
End Sub
Private Sub Command23_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "Videodisc"
End Sub
Private Sub Command24_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "DigitalVideo"
End Sub
Private Sub Command25_Click()
Text22.BackColor = "16711680"
Text21.BackColor = "255"
On Error Resume Next
MMControl1.Command = "Open"
On Error Resume Next
 MMControl1.DeviceType = "MMMovie"
End Sub
Private Sub Command6_Click()
Dim retval
On Error Resume Next
retval = Shell("c:\windows\sndvol32", 1)
On Error Resume Next
retval = Shell("c:\windows\system32\sndvol32", 1)
End Sub
Private Sub Command7_Click()
On Error Resume Next
    dsBuffer.Stop
    Set dsBuffer = Nothing
    
    On Error Resume Next
    tmrwelcome.Enabled = False
domusicstop notep

    midiOutClose (hmidi)
    MMControl1.Command = "Close"

Unload Me
End Sub
Private Sub Command8_Click()

Text13.BackColor = "16711680"
Text14.BackColor = "255"

On Error Resume Next
MMControl1.Command = "Record"
End Sub
Private Sub Command9_Click()

Text14.BackColor = "16711680"
Text13.BackColor = "255"

On Error Resume Next
MMControl1.Command = "Play"
End Sub
'Initialize Waveform and Frequency Selectors - Note this
'can also be done in the Properties Editor. I did it this
'way to allow seeing all the initial settings at a glance.
Private Sub Form_Load()

oncol = 255
offcol = 16777215
    Text8.BackColor = "255"
    Text10.BackColor = "255"
    'Initialize the Waveform Selection option
    
    optSine.Value = True
    
    'Initialize the Freqency Selector buttons
    
    scrDecimal.Max = 0
    scrDecimal.Min = 9
    scrDecimal.Value = 0
    txtDecimal.Enabled = False
    txtDecimal.Text = "." & Str(scrDecimal.Value / 10)
    scrUnits.Max = 0
    scrUnits.Min = 9
    scrUnits.Value = 0
    txtUnits.Enabled = False
    txtUnits.Text = Str(scrUnits.Value)
    scrTens.Max = 0
    scrTens.Min = 9
    scrTens.Value = 4
    txtTens.Enabled = False
    txtTens.Text = Str(scrTens.Value)
    scrHundreds.Max = 0
    scrHundreds.Min = 9
    scrHundreds.Value = 4
    txtHundreds.Enabled = False
    txtHundreds.Text = Str(scrHundreds.Value)
    scrThousands.Max = 0
    scrThousands.Min = 9
    scrThousands.Value = 0
    txtThousands.Enabled = False
    txtThousands.Text = ""
    
    stur = 2
    
    Me.Show
    On Local Error Resume Next
    Set ds = dx.DirectSoundCreate("")
    If Err.Number <> 0 Then
        MsgBox "Unable to start DirectSound"
        End
    End If
    ds.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
    
    'Set the default startup frequency
    conFreq
        
Dim rc As Long
Dim curDevice As Long
On Error Resume Next
    midiOutClose (hmidi)
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If (rc <> 0) Then
        MsgBox "Couldn't open midi device - Error #" & rc
    End If
    baseNote = 35
    channel = 1
    volume = 120
    sldInst.Value = 14
    sldVol.Value = 120
    sldPitch.Value = 50
    Text3.Text = sldInst.Value
    Text4.Text = sldVol.Value
    Text5.Text = sldPitch.Value
    Text16.Text = channel
    
      ' Set properties needed by MCI to open.
      On Error Resume Next
   MMControl1.Notify = False
   MMControl1.Wait = True
   MMControl1.Shareable = False
   MMControl1.DeviceType = "WaveAudio"
       
End Sub
'Cleanup on program exit
Private Sub cmdExit_Click()
On Error Resume Next
    Cleanup
    Unload Me
    
End Sub
'Dispose of the DirectSound Object and its buffer
Private Sub Cleanup()
On Error Resume Next
    If Not (dsBuffer Is Nothing) Then dsBuffer.Stop
    Set dsBuffer = Nothing
    Set ds = Nothing
    Set dx = Nothing
    
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Text1.Text = HScroll1.Value
frequency = HScroll1.Value

    dsBuffer.Stop
    Set dsBuffer = Nothing
    
    If optSine.Value = True Then
        sineWave
    End If
        If optSquare.Value = True Then
        squareWave
    End If
        If optTriangle.Value = True Then
        triangleWave
    End If
        If optSaw.Value = True Then
        sawWave
    End If
        If optNoise = True Then
        noise
    End If
End Sub
Private Sub op1_Click()
If optSine.Value = True Then
        sineWave
    End If
    
    If optSquare.Value = True Then
        squareWave
    End If
    
    If optTriangle.Value = True Then
        triangleWave
    End If
    
    If optSaw.Value = True Then
        sawWave
    End If
    
    If optNoise = True Then
        noise
    End If
End Sub
Private Sub op2_Click()
On Error Resume Next
    dsBuffer.Stop
    Set dsBuffer = Nothing
End Sub
Private Sub op3_Click()
Dim rc As Long
Dim curDevice As Long
On Error Resume Next
midiOutClose (hmidi)
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If (rc <> 0) Then
        MsgBox "Couldn't open midi device - Error #" & rc
    End If
    baseNote = 23
    channel = 15
    volume = 127
    sldInst.Value = 12
    sldVol.Value = 127
    
End Sub
Private Sub op4_Click()
On Error Resume Next
 midiOutClose (hmidi)
End Sub
Private Sub op5_Click()
On Error Resume Next
MMControl1.Command = "Open"
End Sub
Private Sub op6_Click()
On Error Resume Next
MMControl1.Command = "Close"
Text21.BackColor = "16711680"
Text22.BackColor = "255"
End Sub
'Select Decimal and insert decimal point
Private Sub scrDecimal_Change()

    txtDecimal.Text = Str(scrDecimal.Value)
    On Error Resume Next
    conFreq
    
End Sub
'Select Units Value
Private Sub scrUnits_Change()

    txtUnits.Text = Str(scrUnits.Value)
    On Error Resume Next
    conFreq
    
End Sub
'Select Tens Value
Private Sub scrTens_Change()

    txtTens.Text = Str(scrTens.Value)
    On Error Resume Next
    conFreq
    
End Sub
'Select Hundreds and mask leading zeroes
Private Sub scrHundreds_Change()

    txtHundreds.Text = Str(scrHundreds.Value)
    If scrHundreds.Value = 0 And scrThousands.Value = 0 Then
        txtHundreds.Text = ""
    End If
On Error Resume Next
    conFreq
    
End Sub
'Select Thousands and mask leading zero
Private Sub scrThousands_Change()

    txtThousands.Text = Str(scrThousands.Value)
    If scrThousands.Value = 0 Then
        txtThousands.Text = ""
    End If
    On Error Resume Next
    conFreq
    
End Sub
'Concatenate the frequency selector settings
Private Sub conFreq()

    frequency = (scrThousands.Value * 1000) + (scrHundreds.Value * 100) + (scrTens.Value * 10) + scrUnits.Value + (scrDecimal.Value / 10)
    Text2.Text = frequency
        If frequency < 20 Then
            MsgBox "Frequency cannot be lower than 20 Hz."
            frequency = 20
            txtTens.Text = "2"
            txtUnits.Text = "0"
            txtDecimal.Text = ".0"
        End If
        
    End Sub
Private Sub cmdGenerate_Click()

Text8.BackColor = "16711680"
Text7.BackColor = "255"

    If optSine.Value = True Then
        sineWave
    End If
    
    If optSquare.Value = True Then
        squareWave
    End If
    
    If optTriangle.Value = True Then
        triangleWave
    End If
    
    If optSaw.Value = True Then
        sawWave
    End If
    
    If optNoise = True Then
        noise
    End If
    
End Sub
Private Sub sineWave()

    makeFile                                            'Create the file and write header
    On Error Resume Next
    bufferptr = 45                                      'Offset to beginning of waveform
        increment = Pi / (sampleRate / frequency)
        For inputValue = 0 To (2 * Pi) Step increment   'Step around the circle
            sample = Int(amplitude * Sin(inputValue))   'Calculate the sample value
            Put #1, bufferptr, sample                   'Write sample value to file
            bufferptr = bufferptr + 1                   'Increment buffer pointer
        Next inputValue                                 'Loop to the next sample
        
    closeFile                                           'Fill in the rest of the file data
                                                        'and close the file.
    
End Sub
Private Sub squareWave()

    makeFile
    On Error Resume Next
    bufferptr = 45
    period = (sampleRate / frequency)
    state = 1
        If state = 1 Then                   'Positive half cycle
            For inputValue = 0 To period
                sample = amplitude * state
                Put #1, bufferptr, sample
                bufferptr = bufferptr + 1
            Next inputValue
        End If
        
     state = -1
        If state = -1 Then                  'Negative half cycle
            For inputValue = 0 To period
                sample = amplitude * state
                Put #1, bufferptr, sample
                bufferptr = bufferptr + 1
            Next inputValue
        End If
        
    closeFile
    
End Sub
Private Sub sawWave()

    makeFile
    On Error Resume Next
    bufferptr = 45
        period = sampleRate / (frequency / 2)
        For inputValue = 0 To period
            sample = Int(2 * amplitude * (inputValue / period))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
    closeFile
    
End Sub
Private Sub triangleWave()

    makeFile
    On Error Resume Next
    state = 0
    bufferptr = 45
        period = sampleRate / frequency
        If state = 0 Then
        For inputValue = 0 To period / 2    'Generate Positive Slope
            sample = Int(2 * amplitude * (inputValue / period))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
        state = 1
        End If
        If state = 1 Then
        For inputValue = 0 To period        'Generate Negative Slope
            sample = Int((amplitude - 2 * amplitude) - 2 * amplitude * (inputValue - period) / period)
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
        state = 2
        End If
        If state = 2 Then
        For inputValue = 0 To period / 2    'Positive Slope to finish cycle
            sample = Int(amplitude + (2 * amplitude * (inputValue / period)))
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        End If
        
    closeFile
    
End Sub
Private Sub noise()

    Randomize                               'Seed random # generator
    
    makeFile
    
    bufferptr = 45
    period = sampleRate
        For inputValue = 0 To period        'Create 44,100 random samples
            sample = Rnd(amplitude) * 254
            Put #1, bufferptr, sample
            bufferptr = bufferptr + 1
        Next inputValue
        
    closeFile
    
End Sub
'Create the .wav file and write header data
Private Sub makeFile()

    fileName = App.Path & "\temp.wav"
    On Error Resume Next
    Kill fileName                   'REM this line if file does not exist
    On Error Resume Next
    Open fileName For Binary Access Write As #1
        Put #1, 1, "RIFF"           '"RIFF" header
        Put #1, 5, CInt(0)          'Filesize - 8, will write later
        Put #1, 9, "WAVEfmt "       '"WAVEfmt " header - not space after fmt
        Put #1, 17, CLng(16)        'Lenth of format data
        Put #1, 21, CInt(1)         'Wave type PCM
        Put #1, 23, CInt(1)         '1 channel
        Put #1, 25, CLng(44100)     '44.1 kHz SampleRate
        Put #1, 29, CLng(88200)     '(SampleRate * BitsPerSample * Channels) / 8
        Put #1, 33, CInt(2)         '(BitsPerSample * Channels) / 8
        Put #1, 35, CInt(16)        'BitsPerSample
        Put #1, 37, "data"          '"data" Chunkheader
        Put #1, 41, CInt(0)         'Filesize - 44, will write later

End Sub
'Get the file length, write it into the header and close the file.
Private Sub closeFile()

    fileSize = LOF(1)
    Put #1, 5, CLng(fileSize - 8)
    Put #1, 41, CLng(fileSize - 44)
    Close #1
    
    Play
    
End Sub
'Define the DirectSound8 buffer, create it and set the play mode
Private Sub Play()

    Dim bufferDesc As DSBUFFERDESC
    bufferDesc.lFlags = DSBCAPS_STATIC Or DSBCAPS_STICKYFOCUS
    fileName = App.Path & "\temp.wav"
    Set dsBuffer = ds.CreateSoundBufferFromFile(fileName, bufferDesc)
    dsBuffer.Play DSBPLAY_LOOPING
    
End Sub
'Stop playing and clear the DirectSound8 buffer
Private Sub cmdStop_Click()

Text7.BackColor = "16711680"
Text8.BackColor = "255"

On Error Resume Next
    dsBuffer.Stop
    Set dsBuffer = Nothing
    
End Sub
'Change the instrument
Private Sub sldInst_Change()
Dim midimsg As Long
    midimsg = (sldInst.Value * 256) + &HC0 + channel
    Text3.Text = sldInst.Value
    midiOutShortMsg hmidi, midimsg
End Sub
Private Sub sldPitch_Change()
baseNote = sldPitch.Value - 1
    Text4.Text = sldPitch.Value
'notep = note.Value
'Text4.Text = notep
End Sub
Private Sub sldVol_Change()
    volume = sldVol.Value
    Text5.Text = sldVol.Value
End Sub
'=========================================
Private Sub Command1_Click()

Text10.BackColor = "16711680"
Text9.BackColor = "255"

 domusic notep
End Sub
Private Sub Command3_Click()

Text9.BackColor = "16711680"
Text10.BackColor = "255"

domusicstop notep
End Sub
Private Sub Command5_Click()
 Text18.BackColor = "16711680"
 Text19.BackColor = "255"
tmrwelcome.Enabled = False
domusicstop notep

End Sub
Private Sub sw1_Click()
sldInst.Value = 14
Text20.BackColor = "255"
Call swcall(19)
'example below
    'Private Sub Command41_Click()
    'Text20.BackColor = "255"
    'Call swcall(19)
    'End Sub
    'or: if x=5 then call swcall(19)(19 is the note)
End Sub
Private Sub sw2_Click()
sldInst.Value = 10
Text20.BackColor = "255"
Call swcall(19)
End Sub
Private Sub sw3_Click()
sldInst.Value = 70
Text20.BackColor = "255"
Call swcall(19)
End Sub
Private Sub sw4_Click()
sldInst.Value = 80
Text20.BackColor = "255"
Call swcall(19)
End Sub
Private Sub sw5_Click()
Text20.BackColor = "16711680"
End Sub
Private Sub tmrWelcome_Timer() 'plays song
    Static pdemo As Long
    domusicstop pdemo - 7
    If pdemo > 64 Then
        pdemo = 0
        tmrwelcome.Enabled = False
        Exit Sub
    End If
    domusic pdemo + 5
    pdemo = pdemo + 12

   End Sub
 Private Sub Command4_Click()
 Text19.BackColor = "16711680"
 Text18.BackColor = "255"
 tmrwelcome.Enabled = True
End Sub
Private Sub lblTitle_Click(Index As Integer)
    tmrwelcome.Enabled = True
End Sub
'Toggle record on/off
Private Sub cmdRec_Click()

Text12.BackColor = "16711680"
Text11.BackColor = "255"

On Error Resume Next
    If tmrplayback.Enabled Then cmdPlay_Click
    If tmrrec.Enabled Then
        'stop recording
        tmrrec.Enabled = False
        cmdrec.BackColor = &HC0C0FF
        Text11.BackColor = "16711680"

    Else
        'start recording
        rec = ""
        timers = 0
        tmrrec.Enabled = True
        cmdrec.BackColor = &H4040FF
        Text11.BackColor = "255"
        
    End If
End Sub
'Toggle playback on/off
Private Sub cmdPlay_Click()

Text11.BackColor = "16711680"
Text12.BackColor = "255"

Dim X As Long
On Error Resume Next
    If tmrrec.Enabled Then cmdRec_Click
    If tmrplayback.Enabled Then
        'stop playback
        tmrplayback.Enabled = False
        cmdplay.BackColor = &H80FF80
        Text12.BackColor = "16711680"
        For X = 1 To 71 '(stop all notes)
            domusicstop X
        Next X
    Else
        'start playback
        Playin = Split(rec, " ")
        playinc = 0
        tmrplayback.Interval = 10
        tmrplayback.Enabled = True
        cmdplay.BackColor = &HC000&
        Text12.BackColor = "255"
    End If
End Sub
'Save recorded music to disk
Private Sub cmdSave_Click()
Dim ff As Integer
On Error Resume Next

    If tmrrec.Enabled Then cmdRec_Click
    If tmrplayback.Enabled Then cmdPlay_Click
    CommonDialog1.ShowSave
    If Not CommonDialog1.fileName = "" Then
        ff = FreeFile
        Open CommonDialog1.fileName For Binary Access Write As #ff
        Put #ff, , rec
        Close #ff
    End If
End Sub
'Load music from disk
Private Sub cmdLoad_Click()
Dim ff As Integer
Dim temp As Long
On Error Resume Next
    If tmrrec.Enabled Then cmdRec_Click
    If tmrplayback.Enabled Then cmdPlay_Click
    CommonDialog1.ShowOpen
    ff = FreeFile
    If Not CommonDialog1.fileName = "" Then
        rec = ""
        Open CommonDialog1.fileName For Input As ff
        rec = Input(LOF(ff), ff)
        Close ff
    End If
End Sub
'Play back a recording
Private Sub tmrPlayBack_Timer()
Dim getnote() As String
Dim temp As Long
    On Error GoTo Errs
    getnote = Split(Playin(playinc), "x")
    temp = getnote(0)
    If temp < 0 Then 'key-up
        domusicstop Abs(temp)
    Else             'key-down
        domusic temp
    End If
    playinc = playinc + 1
    getnote = Split(Playin(playinc), "x")
    temp = getnote(1) * 50
    If temp = 0 Then 'a 0 means another event happens at the exact same time, so do it now!
        tmrPlayBack_Timer
        Exit Sub
    End If
    tmrplayback.Interval = temp
    Exit Sub
Errs:
    cmdPlay_Click
End Sub
'used during recording
Private Sub tmrRec_Timer()
    timers = timers + 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    dsBuffer.Stop
    Set dsBuffer = Nothing
    
    On Error Resume Next
    tmrwelcome.Enabled = False
domusicstop notep

    midiOutClose (hmidi)
    MMControl1.Command = "Close"

End Sub

