VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Tanx"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
   ForeColor       =   &H00000000&
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   700
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   23
      Left            =   6600
      Picture         =   "frmGame.frx":0D2A
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   308
      Top             =   3720
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   22
      Left            =   6600
      Picture         =   "frmGame.frx":14E1
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   307
      Top             =   3720
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   21
      Left            =   6600
      Picture         =   "frmGame.frx":1C7F
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   306
      Top             =   3720
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   23
      Left            =   6600
      Picture         =   "frmGame.frx":2419
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   305
      Top             =   3600
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   22
      Left            =   6600
      Picture         =   "frmGame.frx":2BD0
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   304
      Top             =   3480
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   21
      Left            =   6600
      Picture         =   "frmGame.frx":336E
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   303
      Top             =   3360
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picBullet2M 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2520
      Picture         =   "frmGame.frx":3B08
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   302
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picBullet2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2400
      Picture         =   "frmGame.frx":3E84
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   301
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picTankM8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Index           =   2
      Left            =   3720
      Picture         =   "frmGame.frx":4200
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   300
      Top             =   4560
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox picTankM8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Index           =   1
      Left            =   3720
      Picture         =   "frmGame.frx":4736
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   299
      Top             =   4560
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox picTankM8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Index           =   0
      Left            =   3600
      Picture         =   "frmGame.frx":4C56
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   298
      Top             =   4560
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox picTank8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Index           =   2
      Left            =   3600
      Picture         =   "frmGame.frx":5177
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   297
      Top             =   4560
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox picTank8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Index           =   1
      Left            =   3840
      Picture         =   "frmGame.frx":5913
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   296
      Top             =   4680
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox picTank8 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1410
      Index           =   0
      Left            =   3720
      Picture         =   "frmGame.frx":608B
      ScaleHeight     =   94
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   295
      Top             =   4800
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   71
      Left            =   0
      Picture         =   "frmGame.frx":6807
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   294
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   71
      Left            =   0
      Picture         =   "frmGame.frx":6CE0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   293
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   70
      Left            =   0
      Picture         =   "frmGame.frx":71B9
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   292
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   70
      Left            =   0
      Picture         =   "frmGame.frx":766A
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   291
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1890
      Index           =   69
      Left            =   0
      Picture         =   "frmGame.frx":7B1B
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   290
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1890
      Index           =   69
      Left            =   0
      Picture         =   "frmGame.frx":801D
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   289
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   68
      Left            =   0
      Picture         =   "frmGame.frx":88CB
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   288
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1890
      Index           =   67
      Left            =   120
      Picture         =   "frmGame.frx":8C55
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   287
      Top             =   720
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   68
      Left            =   0
      Picture         =   "frmGame.frx":914D
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   286
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1890
      Index           =   67
      Left            =   0
      Picture         =   "frmGame.frx":94D7
      ScaleHeight     =   126
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   285
      Top             =   0
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Index           =   66
      Left            =   120
      Picture         =   "frmGame.frx":9D8C
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   284
      Top             =   480
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   65
      Left            =   0
      Picture         =   "frmGame.frx":A289
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   283
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   780
      Index           =   66
      Left            =   0
      Picture         =   "frmGame.frx":A6DC
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   67
      TabIndex        =   282
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   65
      Left            =   0
      Picture         =   "frmGame.frx":ABD9
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   281
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   20
      Left            =   7680
      Picture         =   "frmGame.frx":B635
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   280
      Top             =   720
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   19
      Left            =   7080
      Picture         =   "frmGame.frx":BB38
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   279
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   20
      Left            =   7680
      Picture         =   "frmGame.frx":C030
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   278
      Top             =   240
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1680
      Index           =   19
      Left            =   7080
      Picture         =   "frmGame.frx":C533
      ScaleHeight     =   112
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   277
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   64
      Left            =   0
      Picture         =   "frmGame.frx":CA2B
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   276
      Top             =   0
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   765
      Index           =   64
      Left            =   0
      Picture         =   "frmGame.frx":CDFB
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   275
      Top             =   0
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   63
      Left            =   0
      Picture         =   "frmGame.frx":D389
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   274
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   62
      Left            =   0
      Picture         =   "frmGame.frx":D6FC
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   273
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   345
      Index           =   63
      Left            =   0
      Picture         =   "frmGame.frx":DAA8
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   272
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   660
      Index           =   62
      Left            =   0
      Picture         =   "frmGame.frx":DE6F
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   271
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   61
      Left            =   0
      Picture         =   "frmGame.frx":E344
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   270
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   61
      Left            =   960
      Picture         =   "frmGame.frx":E784
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   269
      Top             =   840
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   18
      Left            =   10440
      Picture         =   "frmGame.frx":EB09
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   268
      Top             =   480
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   18
      Left            =   10320
      Picture         =   "frmGame.frx":11902
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   267
      Top             =   120
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   17
      Left            =   9960
      Picture         =   "frmGame.frx":124F0
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   266
      Top             =   240
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   16
      Left            =   10440
      Picture         =   "frmGame.frx":12F70
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   265
      Top             =   360
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   15
      Left            =   10320
      Picture         =   "frmGame.frx":13B51
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   264
      Top             =   480
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   14
      Left            =   10200
      Picture         =   "frmGame.frx":1470C
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   263
      Top             =   240
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   13
      Left            =   10440
      Picture         =   "frmGame.frx":15365
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   262
      Top             =   120
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   12
      Left            =   10200
      Picture         =   "frmGame.frx":16013
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   261
      Top             =   120
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   11
      Left            =   10200
      Picture         =   "frmGame.frx":16D40
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   260
      Top             =   0
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   17
      Left            =   10200
      Picture         =   "frmGame.frx":17B09
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   259
      Top             =   720
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   16
      Left            =   9720
      Picture         =   "frmGame.frx":1A095
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   258
      Top             =   840
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   15
      Left            =   10200
      Picture         =   "frmGame.frx":1CB56
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   257
      Top             =   840
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   14
      Left            =   10080
      Picture         =   "frmGame.frx":1F7F1
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   256
      Top             =   720
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   13
      Left            =   10200
      Picture         =   "frmGame.frx":227B1
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   255
      Top             =   840
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   12
      Left            =   9960
      Picture         =   "frmGame.frx":25986
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   254
      Top             =   600
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4740
      Index           =   11
      Left            =   10080
      Picture         =   "frmGame.frx":28C8C
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   253
      Top             =   480
      Visible         =   0   'False
      Width           =   4260
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Index           =   10
      Left            =   5160
      Picture         =   "frmGame.frx":2C264
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   252
      Top             =   960
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Index           =   10
      Left            =   5160
      Picture         =   "frmGame.frx":2D3FC
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   251
      Top             =   1080
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Index           =   9
      Left            =   5160
      Picture         =   "frmGame.frx":2E594
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   250
      Top             =   960
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Index           =   9
      Left            =   5280
      Picture         =   "frmGame.frx":2F714
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   249
      Top             =   1080
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   20
      Left            =   2760
      Picture         =   "frmGame.frx":30894
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   248
      Top             =   3240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   19
      Left            =   2760
      Picture         =   "frmGame.frx":30DFB
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   247
      Top             =   3240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   18
      Left            =   2760
      Picture         =   "frmGame.frx":31332
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   246
      Top             =   3240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   17
      Left            =   2880
      Picture         =   "frmGame.frx":31899
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   245
      Top             =   3480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   16
      Left            =   2760
      Picture         =   "frmGame.frx":31DFC
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   244
      Top             =   3240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   15
      Left            =   2760
      Picture         =   "frmGame.frx":3233A
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   243
      Top             =   3360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   20
      Left            =   2760
      Picture         =   "frmGame.frx":3289B
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   242
      Top             =   3240
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   19
      Left            =   2880
      Picture         =   "frmGame.frx":32E02
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   241
      Top             =   3240
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   18
      Left            =   2760
      Picture         =   "frmGame.frx":33339
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   240
      Top             =   3360
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   17
      Left            =   2880
      Picture         =   "frmGame.frx":338A0
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   239
      Top             =   3480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   16
      Left            =   3240
      Picture         =   "frmGame.frx":33E03
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   238
      Top             =   3480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   15
      Left            =   2760
      Picture         =   "frmGame.frx":34341
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   237
      Top             =   3480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   60
      Left            =   0
      Picture         =   "frmGame.frx":348A2
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   236
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   59
      Left            =   120
      Picture         =   "frmGame.frx":34BD7
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   235
      Top             =   720
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   58
      Left            =   0
      Picture         =   "frmGame.frx":35038
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   234
      Top             =   600
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   57
      Left            =   0
      Picture         =   "frmGame.frx":3550A
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   233
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   56
      Left            =   0
      Picture         =   "frmGame.frx":35A71
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   232
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   60
      Left            =   0
      Picture         =   "frmGame.frx":35FD2
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   231
      Top             =   0
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   59
      Left            =   0
      Picture         =   "frmGame.frx":36307
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   230
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   58
      Left            =   0
      Picture         =   "frmGame.frx":36768
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   229
      Top             =   0
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   57
      Left            =   0
      Picture         =   "frmGame.frx":36C3A
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   228
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2400
      Index           =   56
      Left            =   0
      Picture         =   "frmGame.frx":371A1
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   227
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   6
      Left            =   10200
      Picture         =   "frmGame.frx":37702
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1900
      TabIndex        =   226
      Top             =   5880
      Visible         =   0   'False
      Width           =   28500
   End
   Begin VB.PictureBox picTankM7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   0
      Left            =   4800
      Picture         =   "frmGame.frx":4CC18
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   176
      TabIndex        =   225
      Top             =   3840
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox picTank7 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   0
      Left            =   3720
      Picture         =   "frmGame.frx":4D277
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   176
      TabIndex        =   224
      Top             =   3840
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   55
      Left            =   240
      Picture         =   "frmGame.frx":4D8D6
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   223
      Top             =   720
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   54
      Left            =   240
      Picture         =   "frmGame.frx":4DCC8
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   222
      Top             =   840
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   55
      Left            =   0
      Picture         =   "frmGame.frx":4E0BA
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   221
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Index           =   54
      Left            =   0
      Picture         =   "frmGame.frx":4E5F6
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   220
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   8
      Left            =   9240
      Picture         =   "frmGame.frx":4EB31
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   219
      Top             =   1080
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   7
      Left            =   9720
      Picture         =   "frmGame.frx":4F47A
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   218
      Top             =   1440
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   8
      Left            =   9120
      Picture         =   "frmGame.frx":4FF14
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   217
      Top             =   -120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   7
      Left            =   9120
      Picture         =   "frmGame.frx":5154D
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   216
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   53
      Left            =   0
      Picture         =   "frmGame.frx":52B88
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   215
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   52
      Left            =   0
      Picture         =   "frmGame.frx":52F04
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   214
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   51
      Left            =   0
      Picture         =   "frmGame.frx":53262
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   213
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   53
      Left            =   0
      Picture         =   "frmGame.frx":535DF
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   212
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   52
      Left            =   0
      Picture         =   "frmGame.frx":5395B
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   211
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   51
      Left            =   0
      Picture         =   "frmGame.frx":53CB9
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   210
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   50
      Left            =   0
      Picture         =   "frmGame.frx":54036
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   598
      TabIndex        =   209
      Top             =   0
      Visible         =   0   'False
      Width           =   8970
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   49
      Left            =   0
      Picture         =   "frmGame.frx":54433
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   324
      TabIndex        =   208
      Top             =   0
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   48
      Left            =   0
      Picture         =   "frmGame.frx":54807
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   207
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   50
      Left            =   0
      Picture         =   "frmGame.frx":54BBE
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   598
      TabIndex        =   206
      Top             =   0
      Visible         =   0   'False
      Width           =   8970
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   49
      Left            =   0
      Picture         =   "frmGame.frx":54FBB
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   324
      TabIndex        =   205
      Top             =   0
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   48
      Left            =   0
      Picture         =   "frmGame.frx":5538F
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   204
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   47
      Left            =   0
      Picture         =   "frmGame.frx":55746
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   203
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Index           =   47
      Left            =   0
      Picture         =   "frmGame.frx":55A7C
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   202
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   46
      Left            =   0
      Picture         =   "frmGame.frx":55DB2
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   201
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   45
      Left            =   0
      Picture         =   "frmGame.frx":5612F
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   200
      Top             =   0
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   46
      Left            =   0
      Picture         =   "frmGame.frx":565CD
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   199
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   45
      Left            =   0
      Picture         =   "frmGame.frx":5694A
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   198
      Top             =   0
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picExplosionM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Index           =   2
      Left            =   9000
      Picture         =   "frmGame.frx":56DE8
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   197
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox picExplosion 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1575
      Index           =   2
      Left            =   9000
      Picture         =   "frmGame.frx":57815
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   196
      Top             =   5400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   14
      Left            =   960
      Picture         =   "frmGame.frx":58242
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   195
      Top             =   3000
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   13
      Left            =   960
      Picture         =   "frmGame.frx":5865C
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   194
      Top             =   3000
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   12
      Left            =   960
      Picture         =   "frmGame.frx":58AB8
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   193
      Top             =   3000
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   11
      Left            =   960
      Picture         =   "frmGame.frx":58F37
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   192
      Top             =   3000
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   14
      Left            =   240
      Picture         =   "frmGame.frx":593D5
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   191
      Top             =   2880
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   13
      Left            =   240
      Picture         =   "frmGame.frx":597EF
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   190
      Top             =   2880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   12
      Left            =   240
      Picture         =   "frmGame.frx":59C4B
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   189
      Top             =   2880
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   11
      Left            =   240
      Picture         =   "frmGame.frx":5A0CA
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   188
      Top             =   2880
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   44
      Left            =   0
      Picture         =   "frmGame.frx":5A568
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   187
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   43
      Left            =   0
      Picture         =   "frmGame.frx":5A8B9
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   186
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   42
      Left            =   0
      Picture         =   "frmGame.frx":5AC28
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   185
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Index           =   44
      Left            =   0
      Picture         =   "frmGame.frx":5AD04
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   184
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Index           =   43
      Left            =   0
      Picture         =   "frmGame.frx":5B077
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   183
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   42
      Left            =   0
      Picture         =   "frmGame.frx":5B45C
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   182
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   41
      Left            =   0
      Picture         =   "frmGame.frx":5B538
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   181
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   40
      Left            =   0
      Picture         =   "frmGame.frx":5B8FA
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   180
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   39
      Left            =   0
      Picture         =   "frmGame.frx":5BCBC
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   179
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   41
      Left            =   0
      Picture         =   "frmGame.frx":5C00D
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   178
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   40
      Left            =   0
      Picture         =   "frmGame.frx":5C467
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   177
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   39
      Left            =   0
      Picture         =   "frmGame.frx":5C8C1
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   176
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   10
      Left            =   1920
      Picture         =   "frmGame.frx":5CC2B
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   175
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   9
      Left            =   1560
      Picture         =   "frmGame.frx":5D011
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   174
      Top             =   1560
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   8
      Left            =   1200
      Picture         =   "frmGame.frx":5D3EE
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   173
      Top             =   1560
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   10
      Left            =   1920
      Picture         =   "frmGame.frx":5D7B0
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   172
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   9
      Left            =   1560
      Picture         =   "frmGame.frx":5DCEF
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   171
      Top             =   2640
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   8
      Left            =   1200
      Picture         =   "frmGame.frx":5E1C3
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   170
      Top             =   2640
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   6
      Left            =   9240
      Picture         =   "frmGame.frx":5E61D
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   169
      Top             =   1200
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2550
      Index           =   6
      Left            =   9120
      Picture         =   "frmGame.frx":5EF56
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   168
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1470
      Index           =   38
      Left            =   1320
      Picture         =   "frmGame.frx":6059B
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   167
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   37
      Left            =   0
      Picture         =   "frmGame.frx":60A24
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   166
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1470
      Index           =   38
      Left            =   0
      Picture         =   "frmGame.frx":60DB2
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   165
      Top             =   0
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   37
      Left            =   0
      Picture         =   "frmGame.frx":613EE
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   164
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   36
      Left            =   0
      Picture         =   "frmGame.frx":6177C
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   163
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   35
      Left            =   0
      Picture         =   "frmGame.frx":61B11
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   162
      Top             =   0
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   34
      Left            =   0
      Picture         =   "frmGame.frx":61EEA
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   161
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   36
      Left            =   0
      Picture         =   "frmGame.frx":6228F
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   160
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   35
      Left            =   0
      Picture         =   "frmGame.frx":626DA
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   159
      Top             =   0
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   34
      Left            =   0
      Picture         =   "frmGame.frx":62AB3
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   158
      Top             =   0
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1725
      Index           =   5
      Left            =   7920
      Picture         =   "frmGame.frx":62EF8
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   142
      TabIndex        =   157
      Top             =   1440
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   7080
      Picture         =   "frmGame.frx":63434
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   156
      Top             =   840
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1725
      Index           =   5
      Left            =   7080
      Picture         =   "frmGame.frx":638B9
      ScaleHeight     =   115
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   142
      TabIndex        =   155
      Top             =   840
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1020
      Index           =   4
      Left            =   7080
      Picture         =   "frmGame.frx":644E1
      ScaleHeight     =   68
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   154
      Top             =   840
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.PictureBox picTankM6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   2
      Left            =   8160
      Picture         =   "frmGame.frx":64DE0
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   153
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picTankM6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   1
      Left            =   8160
      Picture         =   "frmGame.frx":65299
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   152
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picTankM6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   0
      Left            =   8160
      Picture         =   "frmGame.frx":65736
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   151
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picTank6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   2
      Left            =   7080
      Picture         =   "frmGame.frx":65BD4
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   150
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picTank6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   1
      Left            =   7080
      Picture         =   "frmGame.frx":6608D
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   149
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picTank6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   0
      Left            =   7080
      Picture         =   "frmGame.frx":6652A
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   148
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   33
      Left            =   0
      Picture         =   "frmGame.frx":669C8
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   147
      Top             =   0
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   33
      Left            =   0
      Picture         =   "frmGame.frx":66FA3
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   146
      Top             =   0
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   32
      Left            =   0
      Picture         =   "frmGame.frx":6757E
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   145
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   32
      Left            =   0
      Picture         =   "frmGame.frx":67916
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   144
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   1080
      Picture         =   "frmGame.frx":67DCA
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   143
      Top             =   2280
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   6
      Left            =   1920
      Picture         =   "frmGame.frx":6820A
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   142
      Top             =   2880
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   5
      Left            =   1440
      Picture         =   "frmGame.frx":686D4
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   141
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   120
      Picture         =   "frmGame.frx":68C18
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   140
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   7
      Left            =   240
      Picture         =   "frmGame.frx":690CC
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   139
      Top             =   4320
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1110
      Index           =   6
      Left            =   1080
      Picture         =   "frmGame.frx":69451
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   138
      Top             =   3720
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   810
      Index           =   5
      Left            =   480
      Picture         =   "frmGame.frx":69837
      ScaleHeight     =   54
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   137
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Index           =   4
      Left            =   720
      Picture         =   "frmGame.frx":69C21
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   136
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   31
      Left            =   0
      Picture         =   "frmGame.frx":69FB9
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   135
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   30
      Left            =   0
      Picture         =   "frmGame.frx":6A069
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   134
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   31
      Left            =   0
      Picture         =   "frmGame.frx":6A119
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   133
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   30
      Left            =   0
      Picture         =   "frmGame.frx":6A282
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   132
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   5
      Left            =   9120
      Picture         =   "frmGame.frx":6A3EB
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1690
      TabIndex        =   131
      Top             =   6000
      Visible         =   0   'False
      Width           =   25350
   End
   Begin VB.PictureBox picTankM5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1155
      Index           =   1
      Left            =   7560
      Picture         =   "frmGame.frx":7423A
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   130
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picTankM5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1140
      Index           =   0
      Left            =   7560
      Picture         =   "frmGame.frx":74A63
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   129
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picTank5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1155
      Index           =   1
      Left            =   7560
      Picture         =   "frmGame.frx":7522C
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   128
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picTank5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1140
      Index           =   0
      Left            =   7560
      Picture         =   "frmGame.frx":75A55
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   127
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   29
      Left            =   0
      Picture         =   "frmGame.frx":7621E
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   126
      Top             =   0
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   29
      Left            =   0
      Picture         =   "frmGame.frx":7642C
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   125
      Top             =   0
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3270
      Index           =   3
      Left            =   10320
      Picture         =   "frmGame.frx":7663A
      ScaleHeight     =   218
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   124
      Top             =   4320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3270
      Index           =   3
      Left            =   10440
      Picture         =   "frmGame.frx":7761A
      ScaleHeight     =   218
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   123
      Top             =   4320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2700
      Index           =   2
      Left            =   10320
      Picture         =   "frmGame.frx":785FA
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   122
      Top             =   4200
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1965
      Index           =   1
      Left            =   9600
      Picture         =   "frmGame.frx":79255
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   121
      Top             =   4200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picBossM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1965
      Index           =   0
      Left            =   9600
      Picture         =   "frmGame.frx":79AC1
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   120
      Top             =   4200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2700
      Index           =   2
      Left            =   9720
      Picture         =   "frmGame.frx":7A336
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   119
      Top             =   5040
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1965
      Index           =   1
      Left            =   9600
      Picture         =   "frmGame.frx":7AF91
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   118
      Top             =   4320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picBoss 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1965
      Index           =   0
      Left            =   9600
      Picture         =   "frmGame.frx":7B7FD
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   117
      Top             =   4320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   28
      Left            =   0
      Picture         =   "frmGame.frx":7C072
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   116
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   28
      Left            =   0
      Picture         =   "frmGame.frx":7C434
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   115
      Top             =   0
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox picTankM4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Index           =   0
      Left            =   9000
      Picture         =   "frmGame.frx":7C79E
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   114
      Top             =   5400
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.PictureBox picTank4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Index           =   0
      Left            =   9240
      Picture         =   "frmGame.frx":7CD73
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   113
      Top             =   5520
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.PictureBox picTankM3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   3
      Left            =   9000
      Picture         =   "frmGame.frx":7DE09
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   112
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTankM3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   2
      Left            =   9000
      Picture         =   "frmGame.frx":7E1F4
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   111
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTankM3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   1
      Left            =   9000
      Picture         =   "frmGame.frx":7E5CC
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   110
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTankM3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   0
      Left            =   9000
      Picture         =   "frmGame.frx":7E979
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   109
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTank3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   3
      Left            =   9000
      Picture         =   "frmGame.frx":7EB22
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   108
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTank3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   2
      Left            =   9000
      Picture         =   "frmGame.frx":7EF0D
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   107
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTank3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   1
      Left            =   9000
      Picture         =   "frmGame.frx":7F2E5
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   106
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTank3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   0
      Left            =   9000
      Picture         =   "frmGame.frx":7F692
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   105
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTankM2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   2
      Left            =   9000
      Picture         =   "frmGame.frx":7FA35
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   104
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTankM2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   945
      Index           =   1
      Left            =   9000
      Picture         =   "frmGame.frx":8021B
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   103
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTankM2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   945
      Index           =   0
      Left            =   9000
      Picture         =   "frmGame.frx":809F9
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   102
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTank2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   930
      Index           =   2
      Left            =   9120
      Picture         =   "frmGame.frx":811D0
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   101
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTank2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   945
      Index           =   1
      Left            =   9120
      Picture         =   "frmGame.frx":819B6
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   100
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTank2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   945
      Index           =   0
      Left            =   9120
      Picture         =   "frmGame.frx":82194
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   99
      Top             =   5520
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   27
      Left            =   0
      Picture         =   "frmGame.frx":8296B
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   27
      Left            =   0
      Picture         =   "frmGame.frx":82CCB
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   97
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   26
      Left            =   0
      Picture         =   "frmGame.frx":8302B
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Index           =   26
      Left            =   0
      Picture         =   "frmGame.frx":833B2
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   95
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1605
      Index           =   25
      Left            =   0
      Picture         =   "frmGame.frx":83739
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   24
      Left            =   0
      Picture         =   "frmGame.frx":83C30
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   93
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   23
      Left            =   0
      Picture         =   "frmGame.frx":83FC4
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1605
      Index           =   25
      Left            =   0
      Picture         =   "frmGame.frx":8449D
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   91
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   24
      Left            =   0
      Picture         =   "frmGame.frx":84994
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   90
      Top             =   0
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   23
      Left            =   0
      Picture         =   "frmGame.frx":84D28
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   92
      TabIndex        =   89
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.PictureBox picHelicoptorM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   1
      Left            =   5640
      Picture         =   "frmGame.frx":85201
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   88
      Top             =   4440
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picHelicoptorM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   0
      Left            =   6360
      Picture         =   "frmGame.frx":856A1
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   87
      Top             =   5400
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picHelicoptor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   1
      Left            =   6360
      Picture         =   "frmGame.frx":85C7C
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   86
      Top             =   4920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picHelicoptor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   0
      Left            =   6120
      Picture         =   "frmGame.frx":86252
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   85
      Top             =   5400
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picBoostM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   3840
      Picture         =   "frmGame.frx":8682D
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   84
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBoost 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   3840
      Picture         =   "frmGame.frx":86C73
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   83
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   4
      Left            =   9120
      Picture         =   "frmGame.frx":870AF
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1778
      TabIndex        =   82
      Top             =   6000
      Visible         =   0   'False
      Width           =   26670
   End
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   3
      Left            =   9120
      Picture         =   "frmGame.frx":A4BBC
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1515
      TabIndex        =   81
      Top             =   6000
      Visible         =   0   'False
      Width           =   22725
   End
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   2
      Left            =   9120
      Picture         =   "frmGame.frx":B3BB5
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2010
      TabIndex        =   80
      Top             =   6000
      Visible         =   0   'False
      Width           =   30150
   End
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":C399B
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1990
      TabIndex        =   79
      Top             =   3960
      Visible         =   0   'False
      Width           =   29850
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   22
      Left            =   0
      Picture         =   "frmGame.frx":DE488
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   21
      Left            =   0
      Picture         =   "frmGame.frx":DE9D4
      ScaleHeight     =   142
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   77
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   20
      Left            =   0
      Picture         =   "frmGame.frx":DEF03
      ScaleHeight     =   142
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   76
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   19
      Left            =   0
      Picture         =   "frmGame.frx":DF406
      ScaleHeight     =   142
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Index           =   22
      Left            =   0
      Picture         =   "frmGame.frx":DF5BD
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   74
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   21
      Left            =   0
      Picture         =   "frmGame.frx":DFB09
      ScaleHeight     =   142
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   20
      Left            =   0
      Picture         =   "frmGame.frx":E0038
      ScaleHeight     =   142
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   72
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Index           =   19
      Left            =   0
      Picture         =   "frmGame.frx":E053B
      ScaleHeight     =   142
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   71
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   18
      Left            =   720
      Picture         =   "frmGame.frx":E06F2
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   70
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Index           =   17
      Left            =   720
      Picture         =   "frmGame.frx":E0901
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2025
      Index           =   16
      Left            =   720
      Picture         =   "frmGame.frx":E0D5F
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Index           =   18
      Left            =   0
      Picture         =   "frmGame.frx":E11EB
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1770
      Index           =   17
      Left            =   0
      Picture         =   "frmGame.frx":E13FA
      ScaleHeight     =   118
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2025
      Index           =   16
      Left            =   0
      Picture         =   "frmGame.frx":E1A21
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   15
      Left            =   0
      Picture         =   "frmGame.frx":E1FE8
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   64
      Top             =   4920
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   14
      Left            =   0
      Picture         =   "frmGame.frx":E25C3
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   63
      Top             =   4920
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Index           =   13
      Left            =   0
      Picture         =   "frmGame.frx":E2B9E
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   62
      Top             =   5400
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   15
      Left            =   0
      Picture         =   "frmGame.frx":E318C
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   61
      Top             =   5040
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   750
      Index           =   14
      Left            =   0
      Picture         =   "frmGame.frx":E3767
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   60
      Top             =   5040
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1395
      Index           =   13
      Left            =   0
      Picture         =   "frmGame.frx":E3D42
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   59
      Top             =   5040
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   12
      Left            =   0
      Picture         =   "frmGame.frx":E4D59
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   12
      Left            =   0
      Picture         =   "frmGame.frx":E4E09
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4800
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":E4F72
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1995
      TabIndex        =   56
      Top             =   5880
      Visible         =   0   'False
      Width           =   29925
   End
   Begin VB.PictureBox picSplatM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   2280
      Picture         =   "frmGame.frx":1282B5
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   55
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picSplat 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   2280
      Picture         =   "frmGame.frx":12902F
      ScaleHeight     =   102
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   54
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox picExplosionM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1155
      Index           =   1
      Left            =   9840
      Picture         =   "frmGame.frx":129DA9
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   53
      Top             =   5520
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picExplosionM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   0
      Left            =   9720
      Picture         =   "frmGame.frx":12A385
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   52
      Top             =   5640
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picExplosion 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1155
      Index           =   1
      Left            =   9960
      Picture         =   "frmGame.frx":12A982
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   51
      Top             =   5160
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picExplosion 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   0
      Left            =   9840
      Picture         =   "frmGame.frx":12AF5E
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   50
      Top             =   5880
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   3
      Left            =   1440
      Picture         =   "frmGame.frx":12B55B
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   49
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   2
      Left            =   960
      Picture         =   "frmGame.frx":12B638
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   48
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   1
      Left            =   480
      Picture         =   "frmGame.frx":12B799
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   47
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picMissileM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   480
      Picture         =   "frmGame.frx":12B9E4
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1200
      Index           =   3
      Left            =   480
      Picture         =   "frmGame.frx":12BEB8
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   45
      Top             =   2760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   870
      Index           =   2
      Left            =   480
      Picture         =   "frmGame.frx":12C0A4
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   44
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Index           =   1
      Left            =   480
      Picture         =   "frmGame.frx":12C2F8
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   43
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   480
      Picture         =   "frmGame.frx":12C543
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   42
      Top             =   2760
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox picBulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   2760
      Picture         =   "frmGame.frx":12CA17
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   41
      Top             =   4800
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Left            =   2400
      Picture         =   "frmGame.frx":12CD92
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   40
      Top             =   4680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   11
      Left            =   0
      Picture         =   "frmGame.frx":12D10D
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   11
      Left            =   0
      Picture         =   "frmGame.frx":12D5E1
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   10
      Left            =   0
      Picture         =   "frmGame.frx":12DAB5
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   0
      Picture         =   "frmGame.frx":12DE7B
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   10
      Left            =   0
      Picture         =   "frmGame.frx":12E2C0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   9
      Left            =   0
      Picture         =   "frmGame.frx":12E686
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   54
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   8
      Left            =   0
      Picture         =   "frmGame.frx":12EACB
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   0
      Picture         =   "frmGame.frx":12EE73
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   8
      Left            =   0
      Picture         =   "frmGame.frx":12F2CF
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   0
      Picture         =   "frmGame.frx":12F677
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   46
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   6
      Left            =   0
      Picture         =   "frmGame.frx":12FAD3
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   5
      Left            =   0
      Picture         =   "frmGame.frx":12FBAF
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   4
      Left            =   0
      Picture         =   "frmGame.frx":12FC8B
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   795
      Index           =   3
      Left            =   0
      Picture         =   "frmGame.frx":12FD67
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   6
      Left            =   0
      Picture         =   "frmGame.frx":1301AC
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   5
      Left            =   0
      Picture         =   "frmGame.frx":130288
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Index           =   4
      Left            =   0
      Picture         =   "frmGame.frx":130364
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   795
      Index           =   3
      Left            =   0
      Picture         =   "frmGame.frx":130440
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   0
      Picture         =   "frmGame.frx":130885
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   0
      Picture         =   "frmGame.frx":130C1D
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   0
      Picture         =   "frmGame.frx":130FB5
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   0
      Picture         =   "frmGame.frx":131489
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   79
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox picSpriteM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Index           =   0
      Left            =   360
      Picture         =   "frmGame.frx":13195D
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Index           =   0
      Left            =   0
      Picture         =   "frmGame.frx":131E3C
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picHealthM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   2760
      Picture         =   "frmGame.frx":13231B
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   170
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox picHealth 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   2400
      Picture         =   "frmGame.frx":132AAE
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   170
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox picCannonM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   3
      Left            =   9600
      Picture         =   "frmGame.frx":133241
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picCannonM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   2
      Left            =   9480
      Picture         =   "frmGame.frx":133B0B
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picCannonM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   1
      Left            =   9600
      Picture         =   "frmGame.frx":1342E9
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picCannonM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   0
      Left            =   9600
      Picture         =   "frmGame.frx":1349FA
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picCannon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   3
      Left            =   9600
      Picture         =   "frmGame.frx":135030
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picCannon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   2
      Left            =   9600
      Picture         =   "frmGame.frx":1358FA
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picCannon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   1
      Left            =   9600
      Picture         =   "frmGame.frx":1360D8
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picCannon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1305
      Index           =   0
      Left            =   9600
      Picture         =   "frmGame.frx":1367E9
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.PictureBox picTankM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1065
      Index           =   2
      Left            =   4800
      Picture         =   "frmGame.frx":136E1F
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   146
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox picTank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1065
      Index           =   2
      Left            =   2520
      Picture         =   "frmGame.frx":137541
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   146
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.PictureBox picTankM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   1
      Left            =   4800
      Picture         =   "frmGame.frx":137C63
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picTank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   1
      Left            =   2520
      Picture         =   "frmGame.frx":138378
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picTankM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   0
      Left            =   4800
      Picture         =   "frmGame.frx":138A8D
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.PictureBox picTank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   0
      Left            =   2520
      Picture         =   "frmGame.frx":1391A4
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   2205
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tank properties
Dim HOVER_HEIGHT As Long 'How high you can hover
Dim BULLET_COOLDOWN As Long 'How fast you can fire
Dim TANK_SPEED As Long 'How fast your tank moves
Dim TANK_TYPE As Long 'Type of tank you have
Dim bBoost As Boolean 'If you are allowed to boost or not
Dim intBoost As Long 'Amount of boost remaining
Dim BoostOn As Boolean 'If you are boosting right now
Dim AllowCloak As Boolean 'If cloak is allowed or not
Dim OverHeat As Boolean 'Whehter or not you have overheated your boosters
Dim intCloak As Long 'Amount of cloak remaining
Dim bCloak As Boolean 'Whether cloak is on or not
Dim DoubleHover As Boolean 'Whether double hovering is enabled or not
Dim RapidFire As Boolean 'Whehter rapid fire is enabled or not
Dim TankBlink As Long 'Interval for how long the tank blinks and is invincible after getting hit
Dim FLOOR As Long 'Where the area the tank drives on is
Dim intDemo As Long 'Whether or not the game is in demo mode or not
Dim intTick As Long

Dim bPaused As Boolean 'Is game paused

Dim bCloakUp As Boolean 'Did you unengage cloak yet?

Dim bBeam As Boolean 'Whether you're caught in a tractor beam or not

Dim intRedAlert As Long

Dim LevelFinished As Boolean 'Whether the game has finished

Dim intAsh As Long 'How long ash should stay on the screen

Dim Boss As Sprite 'Boss of the level
Dim BossT As Sprite 'Another value used in various bosses
Dim Plane As Sprite 'Plane that attacks you for the battleship boss
Dim intFlame As Long 'Number of flames the first boss has shot at you
Dim intBossShots As Long 'Number of shots the boss has taken

Dim AllRangeMode As Boolean 'Whether you're in all range mode or not
Dim AllRangeBoost As Long 'Initial prop-up in all range mode
Dim strLevelName As String 'Name of the level
Dim intBlackBars As Long 'Height of the letterbox in the beginning of the level
Dim GROUND_COLOR As Long 'Color of the ground

'Gameloop variables
Dim bRunning As Boolean 'Whehter the game loop is running or not
Dim TickDif As Double 'How fast the game runs
Dim LastTick As Currency 'What the last interval was you executed the loop

Dim Explosion(1 To 5) As Sprite 'Explosion images
Dim BG As Sprite 'Background image
Dim intShake As Long 'Form shaking variable
Dim intHP As Long 'Your HP
Dim strBoss As String 'Which boss it is
Dim bBoss As Boolean 'Whether it is a boss level
Dim bossMode As Boolean 'Whether or not you're facing the boss
Dim LevelType As String 'What the level background is (desert, mountain, etc.)
Dim TotalScroll As Long 'How much of the level has been completed so far

'Input variables (whether the following key is pressed or not)
Dim bLeft As Boolean
Dim bRight As Boolean
Dim bUp As Boolean
Dim bDown As Boolean

Dim HoverState As Long 'Whether you're hovering or not
Dim CannonState As Long 'Is the cannon ready to fire?
Dim CannonWait As Long 'Cannon cooldown delay
Dim ScrollSpeed As Long 'How fast the game scrolls
Dim Tank As Sprite 'The tank
Dim Bullet(1 To 10) As Sprite 'Bullets your tank fires
Dim Enemy(1 To 100) As Sprite 'All enemies/powerups on screen

Private Sub Form_Activate()
If bPaused = False Then 'Not paused, activate the game loop
    bRunning = True
    Call Form_Load
    Call GameLoop
End If
End Sub

Private Sub Form_Deactivate()
'Disable key inputs
bLeft = False
bRight = False
bUp = False
bDown = False
End Sub

Private Sub Form_GotFocus()
'Disable key inputs
bLeft = False
bRight = False
bUp = False
bDown = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Handle key inputs from the game
On Error Resume Next
Dim strDemo As String 'Path for getting demo info
strDemo = App.Path & "\demo.ini"
If Demo = True Then 'If the demo is enabled or not
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = vbKeySpace Then 'Return to main screen if enter or space is pressed
        Unload Me
        frmIntro.Show
        frmIntro.timeDemo.Enabled = True
    End If
    Exit Sub
End If

If LevelFinished = True Then 'Don't load a finished level
    Exit Sub
End If
If KeyCode = vbKeyF1 Then 'Cheat to give 100 HP
    intHP = 100
End If
If KeyCode = vbKeyF2 Then 'Cheat to turn on Rapid Fire mode
    RapidFire = True
End If

If bBeam = False Then 'Not caught in the tractor beam of a UFO
    If KeyCode = vbKeyRight Then
        bRight = True
    End If
    If KeyCode = vbKeyLeft Then
        bLeft = True
    End If
    If KeyCode = vbKeyUp Then 'Attempt to hover
        If AllRangeMode = False And TANK_TYPE <> 6 Then 'A regular hover
            If HoverState = 0 Then 'You are not ocurrently hovering
                PlySound ("hover")
                HoverState = 1
            'You are hovering and double hover is enabled
            ElseIf HoverState > HOVER_HEIGHT And HoverState <= (HOVER_HEIGHT + 25) And DoubleHover = True Then
                HoverState = 100
            End If
        ElseIf TANK_TYPE = 6 Then 'You are a sub, you can freely move up
            bUp = True
        ElseIf AllRangeMode = True Then 'You are in All Range Mode, you can freely move up
            bUp = True
        End If
    End If
    If KeyCode = vbKeyDown Then 'Return to the ground
        bDown = True
    End If
Else
    PlySound ("lowarn") 'Not able to move when you're caught in a tractor beam
End If
If KeyCode = vbKeySpace Then 'Fire
    If CannonState = 3 And TANK_TYPE <> 3 Then 'You are able to fire
        MakeBullet
        If TANK_TYPE = 7 Then
            MakeBullet (True) ' Adjust height for a small tank
        End If
    End If
End If
If KeyCode = vbKeyReturn And bBoost = True And OverHeat = False Then 'If you are able to boost
    If HoverState > 0 Or AllRangeMode = True Then 'If you are hovering
        BoostOn = True 'Turn on the boost
    End If
End If
If KeyCode = vbKeyP Then 'Pause/unpause the game
    If bRunning = True Then 'The game is running, so pause
        bRunning = False
        'Create a text that says "Paused"
        Dim strText As String
        Me.ForeColor = COLRED
        Me.FontSize = 20
        strText = "Paused - Press P To Continue"
        Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 175, 200, strText, Len(strText))
        Me.Refresh
        PauseMidi (curMidi) 'Pause the music
        bPaused = True 'The game is currently paused
    Else 'The game is not running, so resume
        bPaused = False 'The game is no longer paused
        ResumeMidi (curMidi) 'Resume the music
        bRunning = True
        Call GameLoop 'Start the gameloop again
    End If
End If
If KeyCode = vbKeyZ And bCloakUp = False Then 'If you are allowed to cloak
    If bCloak = False And intCloak > 76 Then 'If you are not already cloaking / locket out
        bCloak = True 'Turn on cloak
        bCloakUp = True
    Else
        bCloak = False
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'Stop moving in various directions
If KeyCode = vbKeyRight Then
    bRight = False
End If
If KeyCode = vbKeyLeft Then
    bLeft = False
End If
If KeyCode = vbKeyDown Then
    bDown = False
End If
If KeyCode = vbKeyUp Then
    bUp = False
End If
If KeyCode = vbKeyZ Then
    bCloakUp = False
End If
If KeyCode = vbKeyReturn Then 'Stop boosting
    BoostOn = False
    If TANK_TYPE = 2 Then 'Change to a non-boosting tank picture
        Tank.Picture = 2
    End If
End If
If KeyCode = vbKeyUp And AllRangeMode = True Then 'In all range mode you fall when not holding up
    bUp = False
End If

End Sub

Private Sub Form_Load()
'Load the game
On Error Resume Next
Dim strTemp As String

intDemo = 0

Dim strSave As String
strSave = App.Path & "\" & NextLevel 'Level to load

If LevelLoaded = True Then Exit Sub 'Stop if you've already loaded the level


LevelLoaded = True 'You have now loaded the level
LevelFinished = False 'The level is not finished
bBeam = False 'You are not being trapped by a tractor beam
bCloakUp = False 'You are not cloaking right now
TotalScroll = 0 'You've haven't gone anywhere yet
ScrollSpeed = 5 'How fast the map comes
CannonState = 3 'Set cannon to ready to fire
HoverState = 0 'You are not hovering
RapidFire = False 'You do not possess rapid fire
AllRangeMode = False 'You are not in All Range Mode

'Set where the screen is
frmGame.Left = 100
frmGame.Top = 100

'Turn off movement
bUp = False
bLeft = False
bRight = False
bDown = False

'Set "letterbox" effect bars
intBlackBars = -50

'Load levle data from the INI File:

strLevelName = GetFromIni("GEN", "NAME", strSave) 'Name of level

strTemp = GetFromIni("GEN", "DOUBLEHOVER", strSave) 'If you can double hover or not
If strTemp = "1" Then
    DoubleHover = True
Else
    DoubleHover = False
End If
strTemp = GetFromIni("GEN", "BOOST", strSave) 'If you can boost or not
If strTemp = "1" Then
    bBoost = True
Else
    bBoost = False
End If
strTemp = GetFromIni("GEN", "CLOAK", strSave) 'If you can cloak or not
If strTemp = "1" Then
    intCloak = 25
    AllowCloak = True
Else
    intCloak = 0
    AllowCloak = False
End If
strTemp = GetFromIni("GEN", "FIREDELAY", strSave) 'How long it takes to fire
If strTemp = "" Then strTemp = "20"
BULLET_COOLDOWN = CLng(strTemp)
strTemp = GetFromIni("GEN", "TANKTYPE", strSave) 'What type of tank you have
If strTemp = "" Then strTemp = "0"
TANK_TYPE = CLng(strTemp)
Select Case TANK_TYPE
    Case 0
        Tank.Width = picTank(0).Width
        Tank.Height = picTank(0).Height
        HOVER_HEIGHT = 25
        TANK_SPEED = 6
    Case 1
        Tank.Width = picTank2(0).Width
        Tank.Height = picTank2(0).Height
        HOVER_HEIGHT = 17
        TANK_SPEED = 4
    Case 2
        Tank.Width = picTank3(0).Width
        Tank.Height = picTank3(0).Height
        HOVER_HEIGHT = 24
        TANK_SPEED = 9
    Case 3
        Tank.Width = picTank4(0).Width
        Tank.Height = picTank4(0).Height
        HOVER_HEIGHT = 25
        TANK_SPEED = 6
    Case 4
        Tank.Width = picTank5(0).Width
        Tank.Height = picTank5(0).Height
        HOVER_HEIGHT = 25
        TANK_SPEED = 7
    Case 5
        Tank.Width = picTank6(0).Width
        Tank.Height = picTank6(0).Height
        HOVER_HEIGHT = 15
        TANK_SPEED = 5
    Case 6
        Tank.Width = picTank7(0).Width
        Tank.Height = picTank7(0).Height
        HOVER_HEIGHT = 15
        TANK_SPEED = 5
    Case 7
        Tank.Width = picTank8(0).Width
        Tank.Height = picTank8(0).Height
        HOVER_HEIGHT = 22
        TANK_SPEED = 6
End Select
For i = 1 To 10 'Set all bullets to not being fired
    Bullet(i).Visible = False
    Bullet(i).Width = picBullet.ScaleWidth
    Bullet(i).Height = picBullet.ScaleHeight
Next 'i
TickDif = 0.045 'Set the speed of the game loop
intHP = 100 'Set health to 100%

FLOOR = 320 'Where the bottom of the level is

'Load the background image
LevelType = GetFromIni("GEN", "TYPE", strSave)
If LevelType = "Desert" Then
    BG.Picture = 0
    GROUND_COLOR = COLDESERT
ElseIf LevelType = "City" Then
    BG.Picture = 1
    GROUND_COLOR = COLCITY
    FLOOR = 300
ElseIf LevelType = "Volcano" Then
    BG.Picture = 2
    GROUND_COLOR = COLVOLCANO
ElseIf LevelType = "Mars" Then
    BG.Picture = 3
    GROUND_COLOR = COLMARS
ElseIf LevelType = "Snow" Then
    BG.Picture = 4
    GROUND_COLOR = COLSNOW
ElseIf LevelType = "Ocean" Then
    BG.Picture = 5
    GROUND_COLOR = COLWATERB
ElseIf LevelType = "Underwater" Then
    BG.Picture = 6
    GROUND_COLOR = COLWATER
End If
BG.Width = picBg(BG.Picture).ScaleWidth
BG.Height = FLOOR

'Set tank starting position
Tank.X = 0
Tank.Y = FLOOR - Tank.Height
Tank.Picture = 0

'Determine if and what boss is at the end of this level
strBoss = GetFromIni("GEN", "BOSSNAME", strSave)
strTemp = GetFromIni("GEN", "BOSS", strSave)
If strTemp = "1" Then
    bBoss = True
Else
    bBoss = False
End If
bossMode = False 'Not currently in boss mode
Boss.Visible = False 'Boss not currently visible
Boss.Mode = 0 'Set initial boss state
intBossShots = 0 'Boss hasn't fired yet
intFlame = 1 'No flames shot yet

strTemp = GetFromIni("GEN", "WIDTH", strSave) 'Total width of level
If strTemp = "" Then strTemp = "10"
TotalScroll = (Me.ScaleWidth * (1 + CLng(strTemp))) 'Set how long the level is in pixels
'Load the enemies
For i = 1 To 100
    strTemp = GetFromIni("GEN", "SV" & i, strSave) 'If sprite is visible or not
    If strTemp = "1" Then 'Sprite is visible, get attributes
        Enemy(i).Visible = True
        strTemp = GetFromIni("GEN", "ST" & i, strSave) 'Get type of enemy
        Enemy(i).Picture = CLng(strTemp)
        If Enemy(i).Picture = 3 Then 'Don't draw dive bombs initially
            Enemy(i).Draw = False
        Else
            Enemy(i).Draw = True
        End If
        'Get X/Y coordinates of the enemies
        strTemp = GetFromIni("GEN", "SX" & i, strSave)
        Enemy(i).Width = picSprite(Enemy(i).Picture).ScaleWidth
        Enemy(i).Height = picSprite(Enemy(i).Picture).ScaleHeight
        Enemy(i).X = CLng(strTemp)
        strTemp = GetFromIni("GEN", "SY" & i, strSave)
        Enemy(i).Y = CLng(strTemp)
        Enemy(i).Speed = 4 'How fast the sprite moves
        Enemy(i).Mode = 0 'Set initial mode of sprite
        If TANK_TYPE = 0 Then 'Set the damage it does based on tank
            Enemy(i).Damage = 20
        ElseIf TANK_TYPE = 5 Then
            Enemy(i).Damage = 30
        Else
            Enemy(i).Damage = 10
        End If
        'Set how many shots it takes to kill this enemy
        If Enemy(i).Picture <> 56 And Enemy(i).Picture <> 57 Then
            Enemy(i).HP = 1
        Else
            Enemy(i).HP = 2
        End If
        Enemy(i).SpeedY = 0 'How fast it's moving in Y Direction
        Enemy(i).SecondPic = 0 'Set the current frame
        Enemy(i).Invincible = False 'Sprite not invincible
    Else
        Enemy(i).Visible = False
    End If
    If i <= 5 Then 'Set explosions to being invisible
        Explosion(i).Visible = False
        Explosion(i).Mode = 0
        Explosion(i).Picture = 0
    End If
Next 'i
NextLevel = GetFromIni("GEN", "NEXTLVL", strSave) 'Get the next level to load
End Sub
Private Sub DrawSprites()
'Move and draw the enemies
On Error Resume Next
TotalScroll = TotalScroll - 4 'You've gone this far in the map
If TotalScroll < 0 Then 'You've reached the end of the map
    If bBoss = False Then 'End the map if no boss
        LevelFinished = True
        intBlackBars = Me.ScaleHeight - FLOOR
        Exit Sub
    Else 'Load the boss
        If bossMode = False Then
            PlySound ("bosswarning")
            bossMode = True
        End If
    End If
End If
If AllowCloak = True Then 'If you're able to cloak
    If bCloak = True Then 'If you are cloaked
        intCloak = intCloak - 4 'reduce the amount of cloaking you have left
        If intCloak <= 0 Then bCloak = False
    Else
        intCloak = intCloak + 1 'Slowly increase the cloaking you have remaining
        If intCloak >= 86 Then
            intCloak = 86
        End If
    End If
End If
'Move and draw each enemy
For i = 1 To 100
    If Enemy(i).Visible = True Then 'This enemy is not destroyed / non-existant
        If Enemy(i).X >= Me.ScaleWidth Then 'The enemy is not on the screen
            Enemy(i).X = Enemy(i).X - Enemy(i).Speed
        End If
        'If the enemy has illegally collided with the tank
        If TankBlink = 0 And bCloak = False And Enemy(i).Draw = True And Hit(Enemy(i), Tank) = True And Enemy(i).Picture <> 8 And Enemy(i).Picture <> 39 And Enemy(i).Picture <> 48 And Enemy(i).Picture <> 49 And Enemy(i).Picture <> 50 And Enemy(i).Picture <> 28 Then 'Not a power up or gap
            If Enemy(i).Picture = 70 Then 'Beam, trap the tank in a beam
                If bBeam = False Then
                    Enemy(i).Mode = 100
                    bBeam = True
                    PlySound ("beam")
                End If
            ElseIf Enemy(i).Picture = 29 Then 'Crane
                Call CreateExplosion(Enemy(i))
                PlySound ("explo2")
                intShake = 10
                If intHP = 20 Then
                    PlySound ("warning")
                End If
                If intHP = 0 Then
                    PlySound ("lowarn")
                End If
                If intHP < 0 Then
                    Call GameOver
                    Exit Sub
                End If
                TankBlink = 20
            Else 'Any other enemy
                PlySound ("explo2") 'Play a sound
                Call CreateExplosion(Enemy(i)) 'Draw an explosion
                intShake = 10 'Shake the screen
                Enemy(i).Visible = False 'Set the enemy to invisible
                intHP = intHP - Enemy(i).Damage 'Reduce your health
                If intHP = 20 Then 'Play a warning sound
                    PlySound ("warning")
                End If
                If intHP = 0 Then 'Play a dire warning sound
                    PlySound ("lowarn")
                End If
                If intHP < 0 Then 'You are dead
                    Call GameOver
                    Exit Sub
                End If
                TankBlink = 20 'Set invincible time
            End If
        End If
        'Do special things for certain enemies:
        If Enemy(i).Picture >= 48 And Enemy(i).Picture <= 50 Then
            If Tank.X >= Enemy(i).X And Tank.X + Tank.Width <= Enemy(i).X + Enemy(i).Width And Tank.Y + Tank.Height >= FLOOR Then
                Tank.Y = Tank.Y + 6
                bLeft = False
                bRight = False
                bUp = False
                bDown = False
                If Tank.Y + Tank.Height >= 300 Then
                    Call GameOver
                    Exit Sub
                End If
            End If
        End If
        If Enemy(i).Picture = 3 Then 'Dive bomb
            If Enemy(i).Draw = False Then
                If Enemy(i).X <= Tank.X + (Tank.Width) - 10 Then
                    Enemy(i).Draw = True
                    PlySound ("divebomb")
                    Enemy(i).SpeedY = 5
                End If
            Else
                If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                    Enemy(i).Visible = False
                    Call CreateExplosion(Enemy(i))
                    PlySound ("explo2")
                End If
            End If
        End If
        If Enemy(i).X <= Me.ScaleWidth Then 'The enemy is actually on the screen
            Enemy(i).X = Enemy(i).X - Enemy(i).Speed
            Enemy(i).Y = Enemy(i).Y + Enemy(i).SpeedY
            If Enemy(i).X <= 0 - Enemy(i).Width Then 'The enemy is now off the screen, hide it
                Enemy(i).Visible = False
                If Enemy(i).Picture = 28 Then 'All Range Mode power up passed you by
                    bRunning = False
                    MsgBox "You failed to obtain the All Range Mode Powerup, Game Over!"
                    Call GameOver
                    Exit Sub
                End If
            End If
            If Enemy(i).Draw = True Then 'This object is to be drawn
                'Check if the picture has a second frame to use
                If Enemy(i).Picture = 14 Or Enemy(i).Picture = 15 Or Enemy(i).Picture = 33 Then 'Helicoptors
                    Enemy(i).SecondPic = Enemy(i).SecondPic + 1 'Change the frame
                    If Enemy(i).SecondPic > 1 Then Enemy(i).SecondPic = 0
                    'Update the height and width attributes of the sprite
                    Enemy(i).Width = picHelicoptor(Enemy(i).SecondPic).ScaleWidth
                    Enemy(i).Height = picHelicoptor(Enemy(i).SecondPic).ScaleHeight
                    Call BitBlt(Me.hdc, Enemy(i).X, Enemy(i).Y, Enemy(i).Width, Enemy(i).Height, picHelicoptorM(Enemy(i).SecondPic).hdc, 0, 0, vbSrcAnd)
                    Call BitBlt(Me.hdc, Enemy(i).X, Enemy(i).Y, Enemy(i).Width, Enemy(i).Height, picHelicoptor(Enemy(i).SecondPic).hdc, 0, 0, vbSrcPaint)
                Else 'Regular sprite to draw
                    If Enemy(i).Picture <> 11 And Enemy(i).Picture <> 32 And Enemy(i).Picture <> 40 And Enemy(i).Picture <> 41 And Enemy(i).Picture <> 45 And Enemy(i).Picture <> 56 And Enemy(i).Picture <> 57 Then 'not missile
                        Call BitBlt(Me.hdc, Enemy(i).X, Enemy(i).Y, Enemy(i).Width, Enemy(i).Height, picSpriteM(Enemy(i).Picture).hdc, 0, 0, vbSrcAnd)
                        Call BitBlt(Me.hdc, Enemy(i).X, Enemy(i).Y, Enemy(i).Width, Enemy(i).Height, picSprite(Enemy(i).Picture).hdc, 0, 0, vbSrcPaint)
                    Else
                        Call BitBlt(Me.hdc, Enemy(i).X, Enemy(i).Y, Enemy(i).Width, Enemy(i).Height, picMissileM(Enemy(i).SecondPic).hdc, 0, 0, vbSrcAnd)
                        Call BitBlt(Me.hdc, Enemy(i).X, Enemy(i).Y, Enemy(i).Width, Enemy(i).Height, picMissile(Enemy(i).SecondPic).hdc, 0, 0, vbSrcPaint)
                    End If
                End If
            End If
        End If
        If Enemy(i).Picture = 1 Then 'Ground missile
            If Enemy(i).X <= Me.ScaleWidth Then
                Enemy(i).Speed = 5
            End If
        End If
        If Enemy(i).Picture = 4 Then 'Air to air rock
            If Enemy(i).Mode = 0 Then
                Enemy(i).SpeedY = 3
                If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                    Enemy(i).Mode = 1
                End If
            ElseIf Enemy(i).Mode = 1 Then
                Enemy(i).SpeedY = -3
            End If
        End If
        If Enemy(i).Picture = 5 Then 'Air to low air rock
            If Enemy(i).Mode = 0 Then
                Enemy(i).SpeedY = 3
                If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                    Enemy(i).Mode = 1
                End If
            ElseIf Enemy(i).Mode = 1 Then
                Enemy(i).SpeedY = -2
            End If
        End If
        If Enemy(i).Picture = 6 Then 'Air to ground rock
            If Enemy(i).Mode = 0 Then
                Enemy(i).SpeedY = 3
                If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                    Enemy(i).SpeedY = 0
                    Enemy(i).Y = FLOOR - Enemy(i).Height
                    Enemy(i).Mode = 1
                End If
            End If
        End If
        If Enemy(i).Picture = 7 Or Enemy(i).Picture = 12 Or Enemy(i).Picture = 30 Or Enemy(i).Picture = 31 Then 'Volcanic rock
            If Enemy(i).X <= Me.ScaleWidth Then
                If Enemy(i).Picture = 7 Then 'Medium range
                    Enemy(i).SpeedY = 3
                    Enemy(i).Speed = 5
                ElseIf Enemy(i).Picture = 12 Then 'Short rnage
                    Enemy(i).SpeedY = 3
                ElseIf Enemy(i).Picture = 30 Then 'Medium range
                    Enemy(i).SpeedY = 2
                Else 'Long range
                    Enemy(i).SpeedY = 1
                    Enemy(i).Speed = 5
                End If
                If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                    Enemy(i).Visible = False
                    Call CreateExplosion(Enemy(i))
                    PlySound ("explo2")
                End If
            End If
        End If
            
        If Enemy(i).Picture = 9 Then 'Arcing bullet
            Enemy(i).SpeedY = 4
            If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                Enemy(i).Visible = False
                Call CreateExplosion(Enemy(i))
            End If
        End If
        If Enemy(i).Picture = 11 Then 'Air to ground missile
            Enemy(i).SecondPic = Enemy(i).Mode
            If Enemy(i).Mode = 0 Then
                If Enemy(i).X <= Tank.X + 50 Then
                    Enemy(i).Mode = 1
                    Enemy(i).Speed = 0
                    PlySound ("divebomb")
                End If
            ElseIf Enemy(i).Mode < 3 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
                Enemy(i).Height = picMissile(Enemy(i).Mode).ScaleHeight
                Enemy(i).Width = picMissile(Enemy(i).Mode).ScaleWidth
            Else
                Enemy(i).Height = picMissile(Enemy(i).Mode).ScaleHeight
                Enemy(i).Width = picMissile(Enemy(i).Mode).ScaleWidth
                Enemy(i).SpeedY = 10
                If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                    Enemy(i).Visible = False
                    PlySound ("explo2")
                    Call CreateExplosion(Enemy(i))
                End If
            End If
        End If
        If Enemy(i).Picture = 14 Then 'Bomber coptor
            If Enemy(i).Mode = 0 Then
                If Enemy(i).X <= Tank.X + Tank.Width - 15 Then
                    PlySound ("divebomb")
                    Enemy(i).Mode = 1
                    Dim intNewSprite As Long
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Width = picSprite(59).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(59).ScaleHeight
                    Enemy(intNewSprite).X = Enemy(i).X + 20
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height
                    Enemy(intNewSprite).Picture = 59
                    Enemy(intNewSprite).Speed = 0
                    Enemy(intNewSprite).Mode = 1
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).SpeedY = 5
                End If
            End If
        End If
        If Enemy(i).Picture = 15 Then 'Suicide coptor
            Enemy(i).Speed = 8
            If Enemy(i).Mode = 0 Then
                Enemy(i).SpeedY = 10
                If Enemy(i).Y + Enemy(i).Height >= 300 Then
                    Enemy(i).Mode = 1
                End If
            ElseIf Enemy(i).Mode = 1 Then
                Enemy(i).SpeedY = -10
                If Enemy(i).Y + Enemy(i).Height <= 170 Then
                    Enemy(i).Mode = 0
                End If
            End If
        End If
        If Enemy(i).Picture = 18 Then 'Falling ice
            If Enemy(i).Mode = 0 Then
                Enemy(i).Draw = False
                If Enemy(i).X <= Tank.X + Tank.Width - 30 Then
                    Enemy(i).Y = 0 - Enemy(i).Height
                    Enemy(i).Draw = True
                    Enemy(i).Speed = 0
                    Enemy(i).SpeedY = 2
                    Enemy(i).Mode = 1
                End If
            ElseIf Enemy(i).Mode = 1 Then
                If Enemy(i).Y >= 0 Then
                    Enemy(i).Mode = 2
                    Enemy(i).SpeedY = 8
                End If
            ElseIf Enemy(i).Mode = 2 Then
                If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                    Enemy(i).Visible = False
                    Call CreateExplosion(Enemy(i))
                    PlySound ("explo2")
                End If
            End If
        End If
        If Enemy(i).Picture = 23 Then 'UFO
            If Enemy(i).Mode = 0 Then
                Enemy(i).Y = Enemy(i).Y + 2
                If Enemy(i).Y >= 75 Then
                    Enemy(i).Mode = 1
                End If
                If Enemy(i).X <= Me.ScaleWidth And Enemy(i).X >= Tank.X Then
                    Enemy(i).Mode = 2
                    PlySound ("laser")
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).X = Enemy(i).X + (Enemy(i).Width / 2)
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height
                    Enemy(intNewSprite).Picture = 27
                    Enemy(intNewSprite).Width = picSprite(27).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(27).ScaleWidth
                    Enemy(intNewSprite).Speed = 7
                    Enemy(intNewSprite).SpeedY = 4
                    Enemy(intNewSprite).Invincible = True
                End If
            ElseIf Enemy(i).Mode = 1 Then
                Enemy(i).Y = Enemy(i).Y - 2
                If Enemy(i).Y <= 25 Then
                    Enemy(i).Mode = 0
                End If
            ElseIf Enemy(i).Mode = 2 Then
                Enemy(i).Y = Enemy(i).Y + 2
                If Enemy(i).Y >= 75 Then
                    Enemy(i).Mode = 1
                End If
            End If
        End If
        If Enemy(i).Picture = 24 Then 'Fire cannon
            Enemy(i).Mode = Enemy(i).Mode + 1
            If Enemy(i).Mode = 40 Then
                Enemy(i).Invincible = True
                Enemy(i).Picture = 25
                Enemy(i).Mode = 0
                Enemy(i).Height = picSprite(25).ScaleHeight
                Enemy(i).Width = picSprite(25).ScaleWidth
                Enemy(i).Y = FLOOR - Enemy(i).Height
                If Enemy(i).X <= Me.ScaleWidth Then
                    PlySound ("fire")
                End If
            End If
        End If
        If Enemy(i).Picture = 25 Then 'Fire cannon firing
            Enemy(i).Mode = Enemy(i).Mode + 1
            If Enemy(i).Mode = 10 Then
                Enemy(i).Invincible = True
                Enemy(i).Picture = 24
                Enemy(i).Mode = 0
                Enemy(i).Height = picSprite(24).ScaleHeight
                Enemy(i).Width = picSprite(24).ScaleWidth
                Enemy(i).Y = FLOOR - Enemy(i).Height
            End If
        End If
        If Enemy(i).Picture = 26 Then 'Shake form (does not actually display)
            If Enemy(i).X <= Me.ScaleWidth + 5 Then
                Enemy(i).Visible = False
                PlySound ("volcano")
                intShake = 300
                If LevelType = "Volcano" Then 'Turn on the ash effect
                    intAsh = 300
                End If
            End If
        End If
        If Enemy(i).Picture = 27 Then 'Laser
            If Enemy(i).Y >= FLOOR - Enemy(i).Height Then
                Enemy(i).Visible = False
                Call CreateExplosion(Enemy(i))
                PlySound ("explo2")
            End If
        End If
        If Enemy(i).Picture = 32 Then 'Air to Ground Skimming Missile
            If Enemy(i).Mode = 0 Then
                Enemy(i).SecondPic = 4
                If Enemy(i).X <= Me.ScaleWidth - 100 Then
                    Enemy(i).Mode = 1
                    Enemy(i).SecondPic = 5
                    Enemy(i).Width = picMissile(5).ScaleWidth
                    Enemy(i).Height = picMissile(5).ScaleHeight
                End If
            ElseIf Enemy(i).Mode = 1 Then
                Enemy(i).SecondPic = 6
                Enemy(i).Mode = 2
                Enemy(i).Speed = 0
                Enemy(i).SpeedY = 6
                Enemy(i).Width = picMissile(6).ScaleWidth
                Enemy(i).Height = picMissile(6).ScaleHeight
            ElseIf Enemy(i).Mode = 2 Then
                If Enemy(i).Y + Enemy(i).Height >= 315 Then
                    Enemy(i).Mode = 3
                    Enemy(i).SecondPic = 5
                    Enemy(i).SpeedY = 0
                    Enemy(i).Width = picMissile(5).ScaleWidth
                    Enemy(i).Height = picMissile(5).ScaleHeight
                End If
            ElseIf Enemy(i).Mode = 3 Then
                Enemy(i).SecondPic = 4
                Enemy(i).Mode = 4
                Enemy(i).Speed = 6
                Enemy(i).Width = picMissile(4).ScaleWidth
                Enemy(i).Height = picMissile(4).ScaleHeight
                Enemy(i).Y = Enemy(i).Y + 35
            ElseIf Enemy(i).Mode = 5 Then 'Fall from coptor
                If Enemy(i).Y + Enemy(i).Height >= 300 Then
                    Enemy(i).SecondPic = 4
                    Enemy(i).Width = picMissile(4).ScaleWidth
                    Enemy(i).Height = picMissile(4).ScaleHeight
                    Enemy(i).SpeedY = 0
                    Enemy(i).Speed = 10
                    PlySound ("rocket")
                    Enemy(i).Mode = 6
                End If
            End If
        End If
        If Enemy(i).Picture = 33 Then 'Misile firing helicoptor
            If Enemy(i).Mode = 0 Then
                If Enemy(i).X <= Me.ScaleWidth - 100 Then
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Picture = 32
                    Enemy(intNewSprite).SecondPic = 7
                    Enemy(intNewSprite).Width = picMissile(7).ScaleWidth
                    Enemy(intNewSprite).Height = picMissile(7).ScaleHeight
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height
                    Enemy(intNewSprite).X = Enemy(i).X + 20
                    Enemy(intNewSprite).Speed = 0
                    Enemy(intNewSprite).SpeedY = 6
                    Enemy(intNewSprite).Mode = 5
                    Enemy(i).Mode = 1
                End If
            End If
        End If
        If Enemy(i).Picture = 35 Then 'Torpedo
            If Enemy(i).Mode = 0 Then
                Enemy(i).SpeedY = 6
                If Enemy(i).Y + Enemy(i).Height >= 300 Then
                    Enemy(i).Mode = 1
                    Enemy(i).Picture = 36
                    Enemy(i).SpeedY = 0
                    Enemy(i).Speed = 8
                    Enemy(i).Width = picSprite(36).ScaleWidth
                    Enemy(i).Height = picSprite(36).ScaleHeight
                    PlySound ("splash")
                End If
            End If
        End If
        If Enemy(i).Picture = 40 Then 'Ground to air missile
            If Enemy(i).Mode = 0 Then
                Enemy(i).SecondPic = 8
                If Enemy(i).X <= Me.ScaleWidth - 50 Then
                    Enemy(i).Mode = 1
                    Enemy(i).SecondPic = 9
                    Enemy(i).SpeedY = -7
                    Enemy(i).Speed = 2
                    PlySound ("fire")
                End If
            ElseIf Enemy(i).Mode = 1 Then
                If Enemy(i).Y <= 75 Then
                    Enemy(i).Mode = 2
                    Enemy(i).SecondPic = 10
                    Enemy(i).SpeedY = -5
                    Enemy(i).Speed = 4
                End If
            ElseIf Enemy(i).Mode = 2 Then
                If Enemy(i).Y <= 25 Then
                    Enemy(i).Mode = 3
                    Enemy(i).SecondPic = 4
                    Enemy(i).SpeedY = 0
                    Enemy(i).Speed = 6
                End If
            ElseIf Enemy(i).Mode = 3 Then
                If Enemy(i).X <= Tank.X + Tank.Width - 15 Then
                    Enemy(i).Mode = 4
                    Enemy(i).SecondPic = 5
                    Enemy(i).SpeedY = 5
                    Enemy(i).Speed = 4
                End If
            ElseIf Enemy(i).Mode = 4 Then
                If Enemy(i).Y >= 75 Then
                    Enemy(i).Mode = 5
                    Enemy(i).SecondPic = 6
                    Enemy(i).SpeedY = 8
                    Enemy(i).Speed = 0
                End If
            ElseIf Enemy(i).Mode = 5 Then
                If Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                    Enemy(i).Visible = False
                    Call CreateExplosion(Enemy(i))
                    PlySound ("explo2")
                End If
            End If
            Enemy(i).Width = picMissile(Enemy(i).SecondPic).ScaleWidth
            Enemy(i).Height = picMissile(Enemy(i).SecondPic).ScaleHeight
        End If
        If Enemy(i).Picture = 41 Then 'Ground to high air missile
            If Enemy(i).Mode = 0 Then
                Enemy(i).SecondPic = 8
                If Enemy(i).X <= Me.ScaleWidth - 75 Then
                    Enemy(i).Mode = 1
                    Enemy(i).SpeedY = -8
                    Enemy(i).SecondPic = 9
                    PlySound ("fire")
                End If
            ElseIf Enemy(i).Mode = 1 Then
                If Enemy(i).Y + Enemy(i).Height <= 0 Then
                    Enemy(i).SpeedY = 0
                    Enemy(i).Mode = 2
                End If
            ElseIf Enemy(i).Mode = 2 Then
                If Enemy(i).X <= Tank.X + 45 Then
                    Enemy(i).SecondPic = 6
                    Enemy(i).SpeedY = 8
                    Enemy(i).Speed = 0
                    Enemy(i).Mode = 3
                    PlySound ("divebomb")
                End If
            ElseIf Enemy(i).Mode = 3 Then
                If Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                    Enemy(i).Visible = False
                    Call CreateExplosion(Enemy(i))
                    PlySound ("explo2")
                End If
            End If
            Enemy(i).Width = picMissile(Enemy(i).SecondPic).ScaleWidth
            Enemy(i).Height = picMissile(Enemy(i).SecondPic).ScaleHeight
        End If
        If Enemy(i).Picture = 42 Then 'Bouncing rock
            If Enemy(i).Mode = 0 Then
                Enemy(i).SpeedY = 6
                If Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                    Enemy(i).Mode = 1
                End If
            ElseIf Enemy(i).Mode = 2 Then
                Enemy(i).SpeedY = 6
                If Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                    Enemy(i).Mode = 1
                End If
                If Enemy(i).Y <= 200 Then
                    Enemy(i).SpeedY = 2
                ElseIf Enemy(i).Y <= 260 Then
                    Enemy(i).SpeedY = 4
                Else
                    Enemy(i).SpeedY = 6
                End If
            ElseIf Enemy(i).Mode = 1 Then
                If Enemy(i).Y <= 200 Then
                    Enemy(i).SpeedY = -2
                ElseIf Enemy(i).Y <= 260 Then
                    Enemy(i).SpeedY = -4
                Else
                    Enemy(i).SpeedY = -6
                End If
                If Enemy(i).Y <= 180 Then
                    Enemy(i).Mode = 2
                End If
            End If
        End If
        If Enemy(i).Picture = 43 Then 'Plasma ball
            If Enemy(i).X <= Me.ScaleWidth Then
                Enemy(i).Mode = Enemy(i).Mode + 1
            End If
            If Enemy(i).Mode >= 80 Then
                Call CreateExplosion(Enemy(i))
                Enemy(i).Visible = False
                PlySound ("explo2")
                For q = 1 To 4
                    intNewSprite = FindNewSprite
                    With Enemy(intNewSprite)
                        .Visible = True
                        .Picture = 44
                        .Width = picSprite(44).ScaleWidth
                        .Height = picSprite(44).ScaleHeight
                        .Damage = 5
                        .Draw = True
                        .Invincible = True
                        .HP = 1
                        .Mode = 0
                        .X = Enemy(i).X
                        .Y = Enemy(i).Y
                        Select Case q
                            Case 1
                                .SpeedY = -4
                                .Speed = 4
                            Case 2
                                .SpeedY = 4
                                .Speed = 4
                            Case 3
                                .SpeedY = 4
                                .Speed = -4
                            Case 4
                                .SpeedY = -4
                                .Speed = -4
                        End Select
                    End With
                Next 'q
            End If
        End If
        If Enemy(i).Picture = 44 Then
            If Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                Enemy(i).Visible = False
            End If
            If Enemy(i).Y + Enemy(i).Height <= 0 Then
                Enemy(i).Visible = False
            End If
        End If
        If Enemy(i).Picture = 45 Then 'Shooting turret
            If Enemy(i).X <= Me.ScaleWidth Then
                If Enemy(i).Mode = 0 Then
                    Enemy(i).SecondPic = 11
                    intNewSprite = FindNewSprite
                    PlySound ("turret")
                    With Enemy(intNewSprite)
                        .Visible = True
                        .Picture = 47
                        .Width = picSprite(47).ScaleWidth
                        .Height = picSprite(47).ScaleHeight
                        .X = Enemy(i).X - .Width
                        .Y = Enemy(i).Y + 8
                        .Draw = True
                        .Speed = 6
                        .SpeedY = 0
                        .Invincible = True
                        .Damage = 10
                    End With
                ElseIf Enemy(i).Mode = 4 Then
                    Enemy(i).SecondPic = 12
                ElseIf Enemy(i).Mode = 8 Then
                    Enemy(i).SecondPic = 13
                ElseIf Enemy(i).Mode = 12 Then
                    Enemy(i).SecondPic = 14
                ElseIf Enemy(i).Mode = 16 Then
                    Enemy(i).SecondPic = 13
                ElseIf Enemy(i).Mode = 20 Then
                    Enemy(i).SecondPic = 12
                ElseIf Enemy(i).Mode = 24 Then
                    Enemy(i).SecondPic = 11
                ElseIf Enemy(i).Mode = 28 Then
                    Enemy(i).SecondPic = 12
                ElseIf Enemy(i).Mode = 32 Then
                    Enemy(i).SecondPic = 13
                ElseIf Enemy(i).Mode = 36 Then
                    Enemy(i).SecondPic = 14
                ElseIf Enemy(i).Mode = 40 Then
                    Enemy(i).SecondPic = 13
                ElseIf Enemy(i).Mode = 44 Then
                    Enemy(i).SecondPic = 12
                End If
                Enemy(i).Width = picMissile(Enemy(i).SecondPic).ScaleWidth
                Enemy(i).Height = picMissile(Enemy(i).SecondPic).ScaleHeight
                Enemy(i).Y = FLOOR - Enemy(i).Height
                Enemy(i).Mode = Enemy(i).Mode + 1
                If Enemy(i).Mode >= 48 Then
                    Enemy(i).Mode = 0
                End If
            End If
        End If
        If Enemy(i).Picture = 51 Then 'Arcing bullet down
            If Enemy(i).Mode = 0 Then
                Enemy(i).Speed = 8
                Enemy(i).SpeedY = 4
                If Enemy(i).Y + Enemy(i).Height >= FLOOR - 25 Then
                    Enemy(i).Picture = 52
                    Enemy(i).Width = picSprite(52).ScaleWidth
                    Enemy(i).Height = picSprite(52).ScaleHeight
                    Enemy(i).Mode = 1
                    Enemy(i).SpeedY = 0
                End If
            End If
        End If
        If Enemy(i).Picture = 52 Then 'Arcing bullet horizontal
            If Enemy(i).Mode = 0 Then
                Enemy(i).Speed = 8
                Enemy(i).SpeedY = 2
                If Enemy(i).Y + Enemy(i).Height >= FLOOR - 25 Then
                    Enemy(i).Speed = 12
                    Enemy(i).SpeedY = 0
                    Enemy(i).Mode = 1
                End If
            ElseIf Enemy(i).Mode = 2 Then
                Enemy(i).Speed = 8
                If Enemy(i).X <= 260 Then
                    Enemy(i).Picture = 51
                    Enemy(i).Mode = 1
                    Enemy(i).SpeedY = 6
                    Enemy(i).Width = picSprite(51).ScaleWidth
                    Enemy(i).Height = picSprite(51).ScaleHeight
                End If
            ElseIf Enemy(i).Mode = 3 Then
                If Enemy(i).X <= 260 Then
                    Enemy(i).Picture = 53
                    Enemy(i).Mode = 4
                    Enemy(i).SpeedY = -6
                    Enemy(i).Width = picSprite(53).ScaleWidth
                    Enemy(i).Height = picSprite(53).ScaleHeight
                End If
            End If
        End If
        If Enemy(i).Picture = 53 Then 'Arcing bullet up
            If Enemy(i).Mode = 0 Then
                Enemy(i).Speed = 8
                Enemy(i).SpeedY = -4
                If Enemy(i).Y <= 120 Then
                    Enemy(i).Picture = 52
                    Enemy(i).Mode = 2
                    Enemy(i).Width = picSprite(52).ScaleWidth
                    Enemy(i).Height = picSprite(52).ScaleHeight
                End If
            End If
        End If
        If Enemy(i).Picture = 54 Then 'Plane left
            If Enemy(i).Mode = 0 Then
                If Enemy(i).X <= Me.ScaleWidth Then
                    Enemy(i).Mode = 1
                    Enemy(i).Speed = 8
                End If
            ElseIf Enemy(i).Mode = 1 Then
                If Enemy(i).X <= Tank.X + (Tank.Width / 2) Then
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).X = Enemy(i).X + 20
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height
                    Enemy(intNewSprite).Picture = 3
                    Enemy(intNewSprite).Width = picSprite(3).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(3).ScaleHeight
                    Enemy(intNewSprite).Speed = 0
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = False
                    Enemy(intNewSprite).SpeedY = 4
                    PlySound ("divebomb")
                    Enemy(i).Mode = 2
                End If
            End If
        End If
        If Enemy(i).Picture = 55 Then 'Plane right
            If Enemy(i).Mode = 0 Then
                If Enemy(i).X <= Me.ScaleWidth Then
                    Enemy(i).Mode = 1
                    Enemy(i).Speed = -8
                    Enemy(i).X = 0 - Enemy(i).Width + 10
                End If
            ElseIf Enemy(i).Mode = 1 Then
                If Enemy(i).X >= Tank.X - 50 Then
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).X = Enemy(i).X + 20
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height
                    Enemy(intNewSprite).Picture = 3
                    Enemy(intNewSprite).Width = picSprite(3).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(3).ScaleHeight
                    Enemy(intNewSprite).Speed = 0
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = False
                    Enemy(intNewSprite).SpeedY = 4
                    PlySound ("divebomb")
                    Enemy(i).Mode = 2
                End If
            ElseIf Enemy(i).Mode = 2 Then
                If Enemy(i).X >= Me.ScaleWidth Then
                    Enemy(i).Visible = False
                End If
            End If
        End If
        If Enemy(i).Picture = 56 Then 'Celing Turret
            If Enemy(i).X <= Me.ScaleWidth Then
                If Enemy(i).Mode = 0 Then
                    Enemy(i).SecondPic = 15
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Picture = 47
                    Enemy(intNewSprite).Width = picSprite(47).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(47).ScaleHeight
                    Enemy(intNewSprite).X = Enemy(i).X - Enemy(intNewSprite).Width
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height - 5
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = True
                    Enemy(intNewSprite).Speed = 8
                    Enemy(intNewSprite).SpeedY = 0
                    Enemy(i).X = Enemy(i).X - 1
                    PlySound ("turret")
                    Enemy(i).Mode = 1
                ElseIf Enemy(i).Mode <= 20 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                ElseIf Enemy(i).Mode = 21 Then
                    Enemy(i).SecondPic = 16
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Picture = 60
                    Enemy(intNewSprite).Width = picSprite(60).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(60).ScaleHeight
                    Enemy(intNewSprite).X = Enemy(i).X + 10
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = True
                    Enemy(intNewSprite).Speed = Enemy(i).Speed
                    Enemy(intNewSprite).SpeedY = 6
                    PlySound ("turret")
                    Enemy(i).Mode = 22
                ElseIf Enemy(i).Mode <= 40 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                ElseIf Enemy(i).Mode = 41 Then
                    Enemy(i).SecondPic = 17
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Picture = 47
                    Enemy(intNewSprite).Width = picSprite(47).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(47).ScaleHeight
                    Enemy(intNewSprite).X = Enemy(i).X + Enemy(i).Width
                    Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height - 5
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = True
                    Enemy(intNewSprite).Speed = -8
                    Enemy(intNewSprite).SpeedY = 0
                    PlySound ("turret")
                    Enemy(i).Mode = 42
                ElseIf Enemy(i).Mode <= 60 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                    If Enemy(i).Mode = 60 Then
                        Enemy(i).SecondPic = 16
                        Enemy(i).X = Enemy(i).X - 2
                        intNewSprite = FindNewSprite
                        Enemy(intNewSprite).Visible = True
                        Enemy(intNewSprite).Picture = 60
                        Enemy(intNewSprite).Width = picSprite(60).ScaleWidth
                        Enemy(intNewSprite).Height = picSprite(60).ScaleHeight
                        Enemy(intNewSprite).X = Enemy(i).X + 10
                        Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height
                        Enemy(intNewSprite).Draw = True
                        Enemy(intNewSprite).Damage = 10
                        Enemy(intNewSprite).HP = 1
                        Enemy(intNewSprite).Invincible = True
                        Enemy(intNewSprite).Speed = Enemy(i).Speed
                        Enemy(intNewSprite).SpeedY = 6
                        PlySound ("turret")
                        Enemy(i).Mode = 61
                    End If
                ElseIf Enemy(i).Mode <= 90 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                    If Enemy(i).Mode = 90 Then
                        Enemy(i).Mode = 0
                    End If
                End If
                Enemy(i).Width = picMissile(Enemy(i).SecondPic).ScaleWidth
                Enemy(i).Height = picMissile(Enemy(i).SecondPic).ScaleHeight
            End If
        End If
        If Enemy(i).Picture = 57 Then 'Floor Turret
            If Enemy(i).X <= Me.ScaleWidth Then
                If Enemy(i).Mode = 0 Then
                    Enemy(i).SecondPic = 18
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Picture = 47
                    Enemy(intNewSprite).Width = picSprite(47).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(47).ScaleHeight
                    Enemy(intNewSprite).X = Enemy(i).X - Enemy(intNewSprite).Width
                    Enemy(intNewSprite).Y = Enemy(i).Y + 5
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = True
                    Enemy(intNewSprite).Speed = 8
                    Enemy(intNewSprite).SpeedY = 0
                    Enemy(i).X = Enemy(i).X + 2
                    PlySound ("turret")
                    Enemy(i).Mode = 1
                ElseIf Enemy(i).Mode <= 20 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                ElseIf Enemy(i).Mode = 21 Then
                    Enemy(i).SecondPic = 19
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Picture = 60
                    Enemy(intNewSprite).Width = picSprite(60).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(60).ScaleHeight
                    Enemy(intNewSprite).X = Enemy(i).X + 10
                    Enemy(intNewSprite).Y = Enemy(i).Y - Enemy(intNewSprite).Height
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = True
                    Enemy(intNewSprite).Speed = Enemy(i).Speed
                    Enemy(intNewSprite).SpeedY = -6
                    PlySound ("turret")
                    Enemy(i).Mode = 22
                ElseIf Enemy(i).Mode <= 40 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                ElseIf Enemy(i).Mode = 41 Then
                    Enemy(i).SecondPic = 20
                    intNewSprite = FindNewSprite
                    Enemy(intNewSprite).Visible = True
                    Enemy(intNewSprite).Picture = 47
                    Enemy(intNewSprite).Width = picSprite(47).ScaleWidth
                    Enemy(intNewSprite).Height = picSprite(47).ScaleHeight
                    Enemy(intNewSprite).X = Enemy(i).X + Enemy(i).Width
                    Enemy(intNewSprite).Y = Enemy(i).Y + 5
                    Enemy(intNewSprite).Draw = True
                    Enemy(intNewSprite).Damage = 10
                    Enemy(intNewSprite).HP = 1
                    Enemy(intNewSprite).Invincible = True
                    Enemy(intNewSprite).Speed = -8
                    Enemy(intNewSprite).SpeedY = 0
                    PlySound ("turret")
                    Enemy(i).Mode = 42
                ElseIf Enemy(i).Mode <= 60 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                    If Enemy(i).Mode = 60 Then
                        Enemy(i).SecondPic = 19
                        intNewSprite = FindNewSprite
                        Enemy(intNewSprite).Visible = True
                        Enemy(intNewSprite).Picture = 60
                        Enemy(intNewSprite).Width = picSprite(60).ScaleWidth
                        Enemy(intNewSprite).Height = picSprite(60).ScaleHeight
                        Enemy(intNewSprite).X = Enemy(i).X + 10
                        Enemy(intNewSprite).Y = Enemy(i).Y - Enemy(intNewSprite).Height
                        Enemy(intNewSprite).Draw = True
                        Enemy(intNewSprite).Damage = 10
                        Enemy(intNewSprite).HP = 1
                        Enemy(intNewSprite).Invincible = True
                        Enemy(intNewSprite).Speed = Enemy(i).Speed
                        Enemy(intNewSprite).SpeedY = -6
                        PlySound ("turret")
                        Enemy(i).Mode = 61
                        Enemy(i).X = Enemy(i).X - 6
                    End If
                ElseIf Enemy(i).Mode <= 90 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                    If Enemy(i).Mode = 90 Then
                        Enemy(i).Mode = 0
                    End If
                End If
                Enemy(i).Width = picMissile(Enemy(i).SecondPic).ScaleWidth
                Enemy(i).Height = picMissile(Enemy(i).SecondPic).ScaleHeight
                Enemy(i).Y = FLOOR - Enemy(i).Height
            End If
        End If
        If Enemy(i).Picture = 58 Then 'Submarine
            If Enemy(i).X <= Me.ScaleWidth Then
                If Enemy(i).Mode = 0 Or Enemy(i).Mode = 2 Then
                    Enemy(i).SpeedY = 4
                    If Enemy(i).Mode = 0 Then
                        intNewSprite = FindNewSprite
                        Enemy(intNewSprite).Visible = True
                        Enemy(intNewSprite).Picture = 52
                        Enemy(intNewSprite).Width = picSprite(52).ScaleWidth
                        Enemy(intNewSprite).Height = picSprite(52).ScaleHeight
                        Enemy(intNewSprite).X = Enemy(i).X - Enemy(intNewSprite).Width
                        Enemy(intNewSprite).Y = Enemy(i).Y + Enemy(i).Height - 5
                        Enemy(intNewSprite).Draw = True
                        Enemy(intNewSprite).Damage = 10
                        Enemy(intNewSprite).HP = 1
                        Enemy(intNewSprite).Invincible = True
                        Enemy(intNewSprite).Speed = 6
                        Enemy(intNewSprite).SpeedY = 0
                        Enemy(intNewSprite).Mode = 9
                        PlySound ("explo2")
                        Enemy(i).Mode = 2
                    End If
                    If Enemy(i).Y + Enemy(i).Height >= 275 Then
                        Enemy(i).Mode = 1
                    End If
                ElseIf Enemy(i).Mode = 1 Then
                    Enemy(i).SpeedY = -4
                    If Enemy(i).Y <= 150 Then
                        Enemy(i).Mode = 0
                    End If
                End If
            End If
        End If
        If Enemy(i).Picture = 59 Then 'Underwater divebomb
            If Enemy(i).Mode = 0 Then
                Enemy(i).Draw = False
                If Enemy(i).X <= Tank.X + 60 Then
                    Enemy(i).Draw = True
                    Enemy(i).Speed = 2
                    Enemy(i).Mode = 1
                    Enemy(i).SpeedY = 5
                    Enemy(i).Y = 0 - Enemy(i).Height
                End If
            ElseIf Enemy(i).Mode <= 10 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
                If Enemy(i).Mode = 10 Then
                    Enemy(i).Speed = 0
                End If
            ElseIf Enemy(i).Mode <= 20 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
                If Enemy(i).Mode = 20 Then
                    Enemy(i).Speed = 4
                    Enemy(i).Mode = 1
                End If
            End If
            If Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                Enemy(i).Visible = False
                Call CreateExplosion(Enemy(i))
                PlySound ("explo2")
            End If
        End If
        If Enemy(i).Picture = 60 Then 'Vertical laser
            If Enemy(i).Y + Enemy(i).Height <= 0 Or Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                Enemy(i).Visible = False
            End If
        End If
        If Enemy(i).Picture = 61 Then 'Heat seeking missile
            If Enemy(i).Mode = 0 Then
                If Enemy(i).Y <= Tank.Y Then
                    Enemy(i).Mode = 1
                    Enemy(i).SecondPic = 5
                    If Enemy(i).Y <= FLOOR - 10 Then
                        Enemy(i).SpeedY = 5
                    Else
                        Enemy(i).SpeedY = 0
                    End If
                Else
                    Enemy(i).Mode = 2
                    Enemy(i).SecondPic = 10
                    If Enemy(i).Y >= 5 Then
                        Enemy(i).SpeedY = -5
                    Else
                        Enemy(i).SpeedY = 0
                    End If
                End If
            ElseIf Enemy(i).Mode = 1 Or Enemy(i).Mode = 2 Then
                If Enemy(i).X <= 300 Then
                    Enemy(i).Mode = 3
                    Enemy(i).SpeedY = 0
                    Enemy(i).SecondPic = 4
                End If
            End If
            Enemy(i).Width = picMissile(Enemy(i).SecondPic).ScaleWidth
            Enemy(i).Height = picMissile(Enemy(i).SecondPic).ScaleHeight
        End If
        If Enemy(i).Picture = 62 Then 'Fireball up
            If Enemy(i).Mode = 0 Then
                If Enemy(i).Y + Enemy(i).Height <= 0 Then
                    Enemy(i).Mode = 1
                End If
            ElseIf Enemy(i).Mode <= 7 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
            Else
                Enemy(i).Visible = False
                Dim intFlame As Long
                intFlame = Int(Rnd * 6)
                For q = 0 To 10
                    If q - intFlame < 0 Or q - intFlame > 2 Then
                        intNewSprite = FindNewSprite
                        Enemy(intNewSprite).Visible = True
                        Enemy(intNewSprite).Picture = 63
                        Enemy(intNewSprite).Width = picSprite(63).ScaleWidth
                        Enemy(intNewSprite).Height = picSprite(63).ScaleHeight
                        Enemy(intNewSprite).X = (q * 70)
                        Enemy(intNewSprite).Y = 0 - Enemy(intNewSprite).Height - (q * 25)
                        Enemy(intNewSprite).Draw = True
                        Enemy(intNewSprite).Damage = 10
                        Enemy(intNewSprite).HP = 1
                        Enemy(intNewSprite).Invincible = True
                        Enemy(intNewSprite).Speed = 0
                        Enemy(intNewSprite).SpeedY = Enemy(i).SecondPic
                        Enemy(intNewSprite).Mode = 0
                    End If
                Next 'q
            End If
        End If
        If Enemy(i).Picture = 63 Then 'Fireball down
            If Enemy(i).Y + Enemy(i).Height >= FLOOR Then
                Enemy(i).Visible = False
                Call CreateExplosion(Enemy(i))
            End If
        End If
        If Enemy(i).Picture = 64 Then 'Big ball of fire
            If Enemy(i).Mode = 0 Then
                If Enemy(i).X <= 250 Then
                    Enemy(i).Mode = 1
                    Enemy(i).SpeedY = -8
                End If
            ElseIf Enemy(i).Mode = 2 Then
                If Enemy(i).X <= 250 Then
                    Enemy(i).Mode = 3
                End If
            ElseIf Enemy(i).Mode = 4 Then
                If Enemy(i).X <= 250 Then
                    Enemy(i).Mode = 5
                    Enemy(i).SpeedY = 8
                End If
            End If
        End If
        If Enemy(i).Picture = 65 Then 'Alien Tank
            If Enemy(i).Mode <= 20 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
                Enemy(i).Y = Enemy(i).Y + 2
            Else
                Enemy(i).Mode = Enemy(i).Mode + 1
                If Enemy(i).Mode >= 40 Then Enemy(i).Mode = 1
                Enemy(i).Y = Enemy(i).Y - 2
            End If
            If Enemy(i).Mode Mod 30 = 0 And Enemy(i).X <= Me.ScaleWidth Then
                intNewSprite = FindNewSprite
                PlySound ("laser")
                With Enemy(intNewSprite)
                    .Visible = True
                    .Draw = True
                    .Damage = 10
                    .Speed = 8
                    .SpeedY = 0
                    .Invincible = True
                    .Mode = 0
                    .Picture = 47
                    .Width = picSprite(47).ScaleWidth
                    .Height = picSprite(47).ScaleHeight
                    .X = Enemy(i).X - .Width
                    intRand = Int(Rnd * 3)
                    If intRand = 0 Then
                        .Y = Enemy(i).Y + 16
                    ElseIf intRand = 1 Then
                        .Y = Enemy(i).Y + 26
                    Else
                        .Y = Enemy(i).Y + 65
                    End If
                End With
            End If
        End If
        If Enemy(i).Picture = 66 Then 'Alien fighter
            If Enemy(i).Mode <= 30 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
                Enemy(i).Y = Enemy(i).Y + 2
            ElseIf Enemy(i).Mode <= 60 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
                Enemy(i).Y = Enemy(i).Y - 2
                If Enemy(i).Mode = 60 Then Enemy(i).Mode = 1
            End If
            If Enemy(i).Mode Mod 40 = 0 And Enemy(i).X <= Me.ScaleWidth Then
                intNewSprite = FindNewSprite
                PlySound ("laser")
                With Enemy(intNewSprite)
                    .Visible = True
                    .Draw = True
                    .Damage = 10
                    .Speed = 8
                    .SpeedY = 0
                    .Invincible = True
                    .Mode = 0
                    .Picture = 47
                    .Width = picSprite(47).ScaleWidth
                    .Height = picSprite(47).ScaleHeight
                    .X = Enemy(i).X - .Width
                    .Y = Enemy(i).Y + 39
                End With
                intNewSprite = FindNewSprite
                With Enemy(intNewSprite)
                    .Visible = True
                    .Draw = True
                    .Damage = 10
                    .Speed = 8
                    .SpeedY = 0
                    .Invincible = True
                    .Mode = 0
                    .Picture = 47
                    .Width = picSprite(47).ScaleWidth
                    .Height = picSprite(47).ScaleHeight
                    .X = Enemy(i).X - .Width
                    .Y = Enemy(i).Y + 42
                End With
            End If
        End If
        If Enemy(i).Picture = 67 Or Enemy(i).Picture = 69 Then 'Alien wall
            Enemy(i).Invincible = True
            Enemy(i).Mode = Enemy(i).Mode + 1
            If Enemy(i).Picture = 67 Then
                Enemy(i).Picture = 69
            Else
                Enemy(i).Picture = 67
            End If
            If Enemy(i).Mode = 40 Then
                Enemy(i).Picture = 68
                Enemy(i).Width = picSprite(68).ScaleWidth
                Enemy(i).Height = picSprite(68).ScaleHeight
                Enemy(i).Mode = 0
            End If
        End If
        If Enemy(i).Picture = 68 Then
            Enemy(i).Mode = Enemy(i).Mode + 1
            If Enemy(i).Mode = 40 Then
                Enemy(i).Picture = 67
                Enemy(i).Width = picSprite(67).ScaleWidth
                Enemy(i).Height = picSprite(67).ScaleHeight
                Enemy(i).Mode = 0
                If Enemy(i).X <= Me.ScaleWidth Then
                    PlySound ("buzz")
                End If
            End If
        End If
        If Enemy(i).Picture = 71 Then 'UFO Beam
            If Enemy(i).Mode <= 100 Then
                Enemy(i).Mode = Enemy(i).Mode + 1
                If Enemy(i).Mode = 100 Then
                    Enemy(i).Mode = 0
                    intNewSprite = FindNewSprite
                    With Enemy(intNewSprite)
                        .Visible = True
                        .Picture = 70
                        .Width = picSprite(70).ScaleWidth
                        .Height = picSprite(70).ScaleHeight
                        .Mode = 25
                        .X = Enemy(i).X + 20
                        .Y = Enemy(i).Y + Enemy(i).Height
                        .Speed = Enemy(i).Speed
                        .Invincible = True
                        .Damage = 0
                        .SpeedY = 0
                        .Draw = True
                    End With
                End If
            End If
        End If
        If Enemy(i).Picture = 70 Then 'Beam
            If Enemy(i).Mode <= 25 Then
                Enemy(i).Mode = Enemy(i).Mode - 1
                If Enemy(i).Mode <= 0 Then
                    Enemy(i).Visible = False
                End If
            ElseIf Enemy(i).Mode >= 100 Then
                Tank.X = Tank.X - Enemy(i).Speed
                If Tank.X <= 0 Then Tank.X = 0
                If Enemy(i).Mode <= 125 Then
                    Enemy(i).Mode = Enemy(i).Mode + 1
                    Tank.Y = Tank.Y - 3
                Else
                    Enemy(i).Mode = Enemy(i).Mode + 1
                    Tank.Y = Tank.Y + 3
                    If Enemy(i).Mode = 150 Then
                        bBeam = False
                        Enemy(i).Visible = False
                    End If
                End If
                If Enemy(i).X <= 5 Then
                    bBeam = False
                End If
            End If
        End If
    End If 'enemy(i).visible = false
Next 'i


End Sub
Private Sub DrawBG()
'Draws and moves the background picture
On Error Resume Next
If LevelFinished = False Then
    BG.X = BG.X - 2
End If
If BG.X + BG.Width <= 0 Then 'Wrap the background picture to the right of the screen again
    BG.X = 0
    Call BitBlt(Me.hdc, 0, 0, Me.ScaleWidth, 320, picBg(BG.Picture).hdc, -BG.X, 0, vbSrcCopy)
ElseIf BG.X + BG.Width < Me.ScaleWidth Then 'Need to draw another section of the background picture
    Call BitBlt(Me.hdc, (BG.X + BG.Width), 0, Me.ScaleWidth - (BG.X + BG.Width), 320, picBg(BG.Picture).hdc, 0, 0, vbSrcCopy)
    Call BitBlt(Me.hdc, 0, 0, (BG.X + BG.Width), 320, picBg(BG.Picture).hdc, -BG.X, 0, vbSrcCopy)
Else
    Call BitBlt(Me.hdc, 0, 0, Me.ScaleWidth, 320, picBg(BG.Picture).hdc, -BG.X, 0, vbSrcCopy)
End If

End Sub
Private Function Hit(Sprite1 As Sprite, Sprite2 As Sprite)
'Detects if the rectangle of two sprites are touching
On Error Resume Next
If Sprite1.X <= Sprite2.X + Sprite2.Width And Sprite1.X + Sprite1.Width >= Sprite2.X And Sprite1.Y <= Sprite2.Y + Sprite2.Height And Sprite1.Y + Sprite1.Height >= Sprite2.Y Then
    Hit = True
Else
    Hit = False
End If
End Function
Private Sub GameLoop()
'Runs the actual game, goes every tickDif intervals (in this case 45 milliseconds)
On Error Resume Next
Dim curFreq As Currency 'How fast this computer runs (64 bits)
Dim curStart As Currency 'Start time (64 bits)
Dim dblResult As Double 'Resulting tick time

QueryPerformanceFrequency curFreq       'Get the timer frequency
Do While bRunning = True 'Loop while bRunning is true
    QueryPerformanceCounter curStart        'Get the start time
    dblResult = (curStart - LastTick) / curFreq   'Calculate the duration (in seconds)
    If dblResult > TickDif Then 'The new result is tickDif greater than that last result
        LastTick = curStart 'Set the lastTick to the current time
        Me.Cls 'Clear the screen
        DrawBG 'Draw the background iamge
        If LevelType = "City" Then
            If AllRangeMode = False Then
                DrawBuildings 'Draw the rectangular buildings below the tank
            End If
        End If
        If LevelFinished = False Then
            DrawSprites 'Draw enemies
        End If
        DrawTank 'Draw your tank
        DrawBullet 'Draw your bullets
        intTick = intTick + 1 'Increase the number of times you've gone through this loop
        If bossMode = True Then 'The boss is enabled, call the respective boss function
            If strBoss = "Tank" Then
                Call TankBoss
            ElseIf strBoss = "Crane" Then
                Call CraneBoss
            ElseIf strBoss = "Plane" Then
                Call PlaneBoss
            ElseIf strBoss = "Boat" Then
                Call BoatBoss
            ElseIf strBoss = "Final" Then
                Call FinalBoss
            ElseIf strBoss = "Base" Then
                Call BaseBoss
            ElseIf strBoss = "Mars" Then
                Call MarsBoss
            End If
        End If
        DrawForeground 'Draw things in front of the background and sprites
        DrawExplo 'Draw explosions
        
        If LevelType = "Snow" Then 'Draw snow effect
            Call DrawSnow(COLSNOW, 2, 400)
        ElseIf LevelType = "Ocean" Then 'Draw water below boat
            DrawWater
        End If
        If intAsh > 0 Then 'Draw volcanic ash effect
            Call DrawSnow(COLBLACK, 1, intAsh)
            intAsh = intAsh - 1
        End If
        If intBlackBars <= Me.ScaleHeight - FLOOR Then
            Call BlackBars 'Draw letterbox bars
        End If
        
        If intHP = 0 Then 'Flash red if at 0 HP
            Call RedAlert
        End If
        
        If Demo = True Then 'Execute the demo if in demo mode
            Call doDemo
        End If
        
        If LevelLoaded = False Then Exit Sub 'Stop running when the level is finished
        Me.Refresh
    End If
    DoEvents
Loop
End Sub

Private Sub Form_LostFocus()
On Error Resume Next
'Disable key inputs
bLeft = False
bRight = False
bUp = False
bDown = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Stop the game
On Error Resume Next
If Demo = False Then
    StopMidi (curMidi) 'Stop the music
Else
    frmIntro.Show 'Show the intro screen
End If
If LevelFinished = False Then 'Not don ethe level
    If curCampaign <> "" Then 'You're in campaign mode
        frmCampaign.Show 'Go back to the campaign window
    Else
        frmIntro.Show 'Go back to the intor screen
    End If
End If
'Stop the game from running
bRunning = False
bPaused = False
End Sub
Private Sub CreateExplosion(Sprite1 As Sprite, Optional BigExplo As Boolean = False)
'Create an explosion image
On Error Resume Next
Dim intExplo As Long 'Explosion number to create
For i = 1 To 5
    If Explosion(i).Visible = False Then 'This explosion is not being used
        intExplo = i
        Exit For
    End If
Next 'i
If intExplo = 0 Then Exit Sub 'No empty explosion found

'Determine what explosion to use
If Sprite1.Picture = 0 Or Sprite1.Picture = 16 Or Sprite1.Picture = 17 Or Sprite1.Picture = 20 Or Sprite1.Picture = 21 Or Sprite1.Picture = 19 Then 'Wall explosion
    Explosion(intExplo).Picture = 1 'Sideways (wall explosion)
Else 'Ground explosion
    Explosion(intExplo).Picture = 0
End If
If BigExplo = True Then
    Explosion(intExplo).Picture = 2 'Big explosion for bosses
End If
'Set the explosion to visible and adjust its properties
Explosion(intExplo).Visible = True
Explosion(intExplo).X = Sprite1.X
Explosion(intExplo).Y = Sprite1.Y
Explosion(intExplo).Width = picExplosion(Explosion(intExplo).Picture).ScaleWidth
Explosion(intExplo).Height = picExplosion(Explosion(intExplo).Picture).ScaleHeight
If Explosion(intExplo).Y >= FLOOR - Explosion(intExplo).Height Then 'Out of the limits of the screen
    Explosion(intExplo).Y = FLOOR - Explosion(intExplo).Height
End If
Explosion(intExplo).Mode = 10
End Sub
Private Sub Shake()
'Move the screen up and down
On Error Resume Next
intShake = intShake - 1 'Reduce how much time is left to shake the form
If intShake Mod 2 = 0 Then 'Every other time move the form down
    frmGame.Top = frmGame.Top + 15
    frmGame.Left = frmGame.Left + 15
Else 'Every other time move the form up
    frmGame.Top = frmGame.Top - 15
    frmGame.Left = frmGame.Left - 15
End If
End Sub
Private Sub LoadNextLVL()
On Error Resume Next
'Stop this level and load a new level
Unload Me
If NextLevel = "end" Then 'End of the game, load the credits
    frmCredits.Show
Else
    'Load the various movies or load the briefing screen for the next levels
    If NextLevel = "mov02" Then
        anType = 2
        frmCinema.Show
    ElseIf NextLevel = "mov03" Then
        anType = 3
        frmCinema.Show
    ElseIf NextLevel = "mov04" Then
        anType = 4
        frmCinema.Show
    ElseIf NextLevel = "mov05" Then
        anType = 5
        frmCinema.Show
    ElseIf NextLevel = "mov06" Then
        anType = 6
        frmCinema.Show
    ElseIf NextLevel = "final" Then
        anType = 7
        frmCinema.Show
    Else
        frmBriefing.Show
    End If
End If
End Sub
Private Function FindNewSprite() As Long
'Returns the index of a sprite not being used
For i = 1 To 100
    If Enemy(i).Visible = False Then
        FindNewSprite = i
        Exit Function
    End If
Next 'i
'Nothing has been found
FindNewSprite = 0
End Function
Private Sub MakeBullet(Optional bVertical As Boolean = False)
'Makes a new bullet, and if bVertical=true, make a bullet that shoots up
On Error Resume Next
Dim newBullet As Long 'Index of new bullet to use
For i = 1 To 10
    If Bullet(i).Visible = False Then
        newBullet = i
        Exit For
    End If
Next 'i
If newBullet > 10 Then Exit Sub 'No new bullet found to be able to draw
Bullet(newBullet).Visible = True 'Set the bullet to being visible
If bVertical = False Then 'A regular bullet that travels horizontally
    'Set the bullet's coordinates
    Bullet(newBullet).X = Tank.X + Tank.Width
    If bLeft = True Then
        Bullet(newBullet).Speed = 6
    ElseIf bRight = True Then
        Bullet(newBullet).Speed = 10
    Else
        Bullet(newBullet).Speed = 8
    End If
    If TANK_TYPE = 6 Then
        Bullet(newBullet).Y = Tank.Y + 2
    ElseIf TANK_TYPE = 5 Then
        Bullet(newBullet).Y = Tank.Y + 7
    ElseIf TANK_TYPE = 4 Then 'Boat
        Bullet(newBullet).Y = Tank.Y + 14
    Else
        Bullet(newBullet).Y = Tank.Y + 20
    End If
    Bullet(newBullet).Picture = 0
    Bullet(newBullet).Width = picBullet.Width
    Bullet(newBullet).Height = picBullet.Height
    Bullet(newBullet).SpeedY = 0
Else 'Vertical bullet, shoots up
    Bullet(newBullet).X = Tank.X + 27
    Bullet(newBullet).Width = picBullet2.ScaleWidth
    Bullet(newBullet).Height = picBullet2.ScaleHeight
    Bullet(newBullet).Y = Tank.Y - Bullet(newBullet).Height
    Bullet(newBullet).Speed = 0
    Bullet(newBullet).SpeedY = 6
    Bullet(newBullet).Picture = 1
End If
CannonWait = 0
CannonState = 0
PlySound ("gun1") 'Play a gun sound
End Sub
Private Sub BlackBars()
'Draw letterbox effect and beginning and end of level
On Error Resume Next
If LevelFinished = False Then 'Beginning of level
    intBlackBars = intBlackBars + 2
Else 'End of level
    intBlackBars = intBlackBars - 2
    If intBlackBars <= -100 Then
        LoadNextLVL
    End If
End If
If intBlackBars <= 0 Then
    Line (0, 0)-Step(Me.ScaleWidth, Me.ScaleHeight - FLOOR), tempColor, BF
    Line (0, FLOOR)-Step(Me.ScaleWidth, Me.ScaleHeight - FLOOR), tempColor, BF
Else
    Line (0, 0)-Step(Me.ScaleWidth, Me.ScaleHeight - FLOOR - intBlackBars), tempColor, BF
    Line (0, FLOOR + intBlackBars)-Step(Me.ScaleWidth, Me.ScaleHeight - FLOOR - intBlackBars), tempColor, BF
End If
'Draw the level name next
Dim strText As String
Me.Font = "Comic Sans MS"
Me.FontBold = True
Me.FontSize = 14
Me.ForeColor = COLWHITE
If LevelFinished = False Then
    strText = strLevelName
Else
    strText = strLevelName & " Completed!"
End If
If Me.ScaleHeight - FLOOR - intBlackBars > 35 Then
    Call TextOut(Me.hdc, 10, Me.ScaleHeight - 35, strText, Len(strText))
End If
End Sub
Private Sub GameOver()
'You are dead
On Error Resume Next
'Create an explosion for your tank
Explosion(3).Width = picExplosion(2).ScaleWidth
Explosion(3).Height = picExplosion(2).ScaleHeight
Explosion(3).Picture = 2
Explosion(3).X = Tank.X
Explosion(3).Y = Tank.Y
If Explosion(3).Y + Explosion(3).Height >= FLOOR Then
    Explosion(3).Y = FLOOR - Explosion(3).Height
End If
PlySound ("explo") 'Play an explosion sound
bRunning = False 'Stop the game from running
MsgBox "Game Over!"
'Unload the level and load the campaign screen
Unload Me
frmCampaign.Show
End Sub
Private Sub TankBoss()
'Controls the Desert Tank Boss
On Error Resume Next
If Boss.Mode = 0 Then
    Boss.Width = picBoss(0).ScaleWidth
    Boss.Height = picBoss(0).ScaleHeight
    Boss.X = Me.ScaleWidth
    Boss.Y = FLOOR - Boss.Height
    Boss.Picture = 0
    Boss.HP = 3
    Boss.Visible = True
    Boss.Mode = 1
    Boss.Speed = 4
ElseIf Boss.Mode = 1 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.Picture = 0 Then
        Boss.Picture = 1
    Else
        Boss.Picture = 0
    End If
    If Boss.X <= Me.ScaleWidth - Boss.Width + 50 Then
        Boss.Mode = 2
        Enemy(1).Picture = 9
        Enemy(1).Width = picSprite(9).ScaleWidth
        Enemy(1).Height = picSprite(9).ScaleHeight
        Enemy(1).X = Boss.X - Enemy(1).Width
        Enemy(1).Y = Boss.Y + 50
        Enemy(1).Speed = 4
        Enemy(1).SpeedY = 4
        Enemy(1).Invincible = True
        Enemy(1).Visible = True
    End If
ElseIf Boss.Mode = 2 Then
    If Enemy(1).Visible = False Then
        intShake = 15
        PlySound ("explo2")
        Boss.Mode = 3
        intFlame = 1
    End If
ElseIf Boss.Mode < 70 Then
    Boss.Mode = Boss.Mode + 1
    If Boss.Mode Mod 5 = 0 Then
        Enemy(intFlame).Visible = True
        Enemy(intFlame).Draw = True
        Enemy(intFlame).Picture = 10
        Enemy(intFlame).Width = picSprite(10).ScaleWidth
        Enemy(intFlame).Height = picSprite(10).ScaleHeight
        Enemy(intFlame).Y = FLOOR - Enemy(intFlame).Height
        Enemy(intFlame).X = Boss.X - (30 * intFlame) - 30
        Enemy(intFlame).Speed = 0
        Enemy(intFlame).SpeedY = 0
        Enemy(intFlame).Damage = 10
        intFlame = intFlame + 1
    End If
ElseIf Boss.Mode < 140 Then
    If Boss.Mode = 70 Then intFlame = 1
    Boss.Mode = Boss.Mode + 1
    If Boss.Mode Mod 5 = 0 Then
        Enemy(intFlame).Visible = False
        intFlame = intFlame + 1
    End If
ElseIf Boss.Mode = 140 Then
    Boss.X = Boss.X + 4
    If Boss.Picture = 0 Then
        Boss.Picture = 1
    Else
        Boss.Picture = 0
    End If
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 201
    End If
ElseIf Boss.Mode < 250 Then
    Boss.Mode = Boss.Mode + 1
    If Boss.Mode = 249 Then
        intBossShots = intBossShots + 1
        If intBossShots >= 2 Then
            Boss.Mode = 250
        Else
            Boss.Mode = 1
        End If
    End If
ElseIf Boss.Mode = 250 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.Picture = 0 Then
        Boss.Picture = 1
    Else
        Boss.Picture = 0
    End If
    If Boss.X <= Me.ScaleWidth - Boss.Width + 50 Then
        Boss.Picture = 2
        Boss.Width = picBoss(2).ScaleWidth
        Boss.Height = picBoss(2).ScaleHeight
        Boss.Y = FLOOR - picBoss(2).Height
        Boss.Mode = 251
    End If
ElseIf Boss.Mode <= 350 Then
    '155 x1, 40 y1, 200 x2, 85 y2
    Boss.Mode = Boss.Mode + 1
    For i = 1 To 10
        If Bullet(i).Visible = True Then
            If Bullet(i).X <= Boss.X + 200 And Bullet(i).X + Bullet(i).Width >= Boss.X + 155 And Bullet(i).Y <= Boss.Y + 85 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 40 Then
                Call CreateExplosion(Bullet(i))
                Bullet(i).Visible = False
                PlySound ("explo2")
                Boss.HP = Boss.HP - 1
                Boss.Speed = Boss.Speed + 2
                If Boss.HP = 0 Then
                    Boss.Mode = 1000
                Else
                    Boss.Mode = 351
                    Boss.Picture = 0
                    Boss.Height = picBoss(0).ScaleHeight
                    Boss.Width = picBoss(0).ScaleWidth
                    Boss.Y = FLOOR - Boss.Height
                End If
            End If
        End If
    Next 'i
    If Boss.Mode = 350 Then 'out of time
        Boss.Picture = 0
        Boss.Height = picBoss(0).ScaleHeight
        Boss.Width = picBoss(0).ScaleWidth
        Boss.Y = FLOOR - Boss.Height
    End If
ElseIf Boss.Mode = 351 Then
    Boss.X = Boss.X + 4
    If Boss.Picture = 0 Then
        Boss.Picture = 1
    Else
        Boss.Picture = 0
    End If
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 1
        intBossShots = 0
    End If
ElseIf Boss.Mode = 1000 Then
    Boss.Visible = False
    LevelFinished = True
    intBlackBars = Me.ScaleHeight - FLOOR
    Boss.Mode = 1001
End If

If Hit(Boss, Tank) = True And TankBlink = 0 And bCloak = False Then
    intHP = intHP - 20
    TankBlink = 20
    Call CreateExplosion(Tank)
    PlySound ("explo2")
    If intHP < 0 Then
        Call GameOver
    End If
End If

If Boss.Visible = True Then
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBossM(Boss.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBoss(Boss.Picture).hdc, 0, 0, vbSrcPaint)
End If

End Sub
Private Sub CraneBoss()
'Handles the movement of the boss with a giant crane that causes an avalanche when it falls
On Error Resume Next
If Boss.Mode = 0 Then
    Boss.Picture = 3
    Boss.Visible = True
    Boss.Width = picBoss(3).ScaleWidth
    Boss.Height = picBoss(3).ScaleHeight
    Boss.X = Me.ScaleWidth
    Boss.Y = FLOOR - Boss.Height
    Boss.Speed = 4
    Boss.Mode = 1
    Boss.HP = 3
    Enemy(1).Picture = 29
    Enemy(1).Visible = True
    Enemy(1).Draw = True
    Enemy(1).Damage = 0
    Enemy(1).Invincible = False
    Enemy(1).Width = picSprite(29).ScaleWidth
    Enemy(1).Height = picSprite(29).ScaleHeight
    Enemy(1).X = Boss.X - 7
    Enemy(1).Y = Boss.Y + 110
    Enemy(1).Speed = Boss.Speed
    Enemy(1).HP = 3
ElseIf Boss.Mode = 1 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.X + Boss.Width <= Me.ScaleWidth + 5 Then
        Boss.Mode = 2
        Enemy(1).Speed = 0
        Enemy(1).SpeedY = 5
    End If
    If Enemy(1).Visible = False Then 'Crane destroyed
        Boss.HP = Boss.HP - 1
        Boss.Mode = 5
    End If
ElseIf Boss.Mode = 2 Then
    If Enemy(1).Y + Enemy(1).Height >= FLOOR Then   'Crane drop
        Call CreateExplosion(Enemy(1), True)
        Enemy(1).Visible = False
        intShake = 30
        PlySound ("explo2")
        Boss.Mode = 3
    End If
    If Enemy(1).HP = 0 Then 'Crane destroyed
        Boss.HP = Boss.HP - 1
        Boss.Mode = 5
    End If
ElseIf Boss.Mode = 3 Then 'Boss move away
    Boss.X = Boss.X + Boss.Speed
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 4
        For i = 2 To 6
            Enemy(i).Visible = True
            Enemy(i).Picture = 18
            Enemy(i).Draw = True
            Enemy(i).Damage = 10
            Enemy(i).Width = picSprite(18).ScaleWidth
            Enemy(i).Height = picSprite(18).ScaleHeight
            Enemy(i).HP = 1
            Enemy(i).Invincible = False
            Randomize
            Enemy(i).X = Int(Rnd * 500)
            Enemy(i).Y = -Int(Rnd * 500)
            Enemy(i).Mode = 2
            Enemy(i).Speed = 0
            Enemy(i).SpeedY = 5
        Next 'i
    End If
ElseIf Boss.Mode = 4 Then 'Wait for icicles to disapear
    Dim bNextMode As Boolean
    bNextMode = True
    For i = 2 To 6
        If Enemy(i).Visible = True Then
            bNextMode = False
            Exit For
        End If
    Next 'i
    If bNextMode = True Then
        Boss.Mode = 1
        Enemy(1).Picture = 29
        Enemy(1).Visible = True
        Enemy(1).Draw = True
        Enemy(1).Damage = 0
        Enemy(1).Invincible = False
        Enemy(1).Width = picSprite(29).ScaleWidth
        Enemy(1).Height = picSprite(29).ScaleHeight
        Enemy(1).X = Boss.X - 7
        Enemy(1).Y = Boss.Y + 110
        Enemy(1).SpeedY = 0
        Enemy(1).Speed = Boss.Speed
    End If
ElseIf Boss.Mode = 5 Then 'Away without the crane
    Boss.X = Boss.X + Boss.Speed
    If Boss.X >= Me.ScaleWidth Then
        Boss.Speed = Boss.Speed + 2
        Enemy(1).Picture = 29
        Enemy(1).Visible = True
        Enemy(1).Draw = True
        Enemy(1).Damage = 0
        Enemy(1).Invincible = False
        Enemy(1).Width = picSprite(29).ScaleWidth
        Enemy(1).Height = picSprite(29).ScaleHeight
        Enemy(1).X = Boss.X - 7
        Enemy(1).Y = Boss.Y + 110
        Enemy(1).Speed = Boss.Speed
        Enemy(1).SpeedY = 0
        Enemy(1).HP = 3
        Boss.Mode = 1
        If Boss.HP <= 0 Then
            Boss.Visible = False
            Boss.Mode = 6
        End If
    End If
ElseIf Boss.Mode = 6 Then 'Win
    PlySound ("explosion")
    Boss.Visible = False
    Enemy(1).Visible = False
    LevelFinished = True
    intBlackBars = Me.ScaleHeight - FLOOR
    Boss.Mode = 7
End If

If Hit(Boss, Tank) = True Then
    If TankBlink = 0 Then
        TankBlink = 30
        Call CreateExplosion(Tank)
        PlySound ("explo2")
        intHP = intHP - 10
        If intHP < 0 Then
            Call GameOver
        End If
    End If
End If
'Draw the boss
If Boss.Visible = True Then
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBossM(Boss.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBoss(Boss.Picture).hdc, 0, 0, vbSrcPaint)
End If

End Sub
Private Sub DrawSnow(intColor As Long, intRadius As Long, intNumber As Long)
'Draw the "snow" which are random white circles in the foreground
On Error Resume Next
Randomize
For i = 1 To intNumber 'Draw up to intNumber of snow flakes
    Me.ForeColor = intColor
    Me.FillColor = intColor
    Me.FillStyle = 0
    Circle (Int(Rnd * Me.ScaleWidth), Int(Rnd * FLOOR)), intRadius
Next 'i
End Sub
Private Sub DrawWater()
'Draw a rectanle representing water in front of the tank
On Error Resume Next
Line (0, 300)-Step(Me.ScaleWidth, 20), COLWATER, BF
End Sub
Private Sub DrawTank()
'Draws your tank and handles its movement
On Error Resume Next
On Error Resume Next
'Handle key input
If bLeft = True Then 'Move to the left
    Tank.X = Tank.X - TANK_SPEED 'Move tank
    If Tank.X < 0 Then Tank.X = 0 'You are past the boundary of the level
End If
If bRight = True Then
    Tank.X = Tank.X + TANK_SPEED 'Move tank
    If TANK_TYPE <> 6 Then 'Check if you're against the regular boundary
        If Tank.X > 450 Then
            Tank.X = 450
            PlySound ("buzzer")
        End If
    Else 'The city tank is able to move further to the right
        If Tank.X > 550 Then
            Tank.X = 550
            PlySound ("buzzer")
        End If
    End If
End If
If bDown = True Then 'Stop the tank from hovering
    If TANK_TYPE = 6 Then 'This tank is in all range mode
        Tank.Y = Tank.Y + TANK_SPEED
        If Tank.Y + Tank.Height >= FLOOR Then
            Tank.Y = FLOOR - Tank.Height
            PlySound ("buzzer")
        End If
    Else 'Regular tank
        If HoverState > 0 And HoverState < 100 Then 'Stop hovering
            HoverState = HOVER_HEIGHT + 25
        ElseIf HoverState >= 100 And HoverState <= 115 Then 'Stop hovering
            HoverState = 116
        End If
    End If
End If
If bUp = True Then
    If TANK_TYPE = 6 Then 'Move the tank up when it can
        Tank.Y = Tank.Y - TANK_SPEED
        If Tank.Y <= 0 Then
            Tank.Y = 0
            PlySound ("buzzer")
        End If
    End If
End If
'Handle tank hovering
If HoverState > 0 And HoverState < HOVER_HEIGHT Then 'Tank is moving up
    'Change the tank's picture
    If TANK_TYPE < 3 Then
        Tank.Picture = 2
    ElseIf TANK_TYPE = 4 Then
        Tank.Picture = 1
    ElseIf TANK_TYPE = 5 Then
        Tank.Picture = 2
    ElseIf TANK_TYPE = 7 Then
        Tank.Picture = 2
    End If
    Tank.Y = Tank.Y - 3 'Mover the tank up
    HoverState = HoverState + 1
ElseIf HoverState >= HOVER_HEIGHT And HoverState < HOVER_HEIGHT + 25 Then 'Tank it at the highest of its arc
    HoverState = HoverState + 1
    Tank.Mode = Tank.Mode + 1
    If Tank.Mode < 10 Then
        Tank.Y = Tank.Y - 1
    Else
        Tank.Y = Tank.Y + 1
        If Tank.Mode >= 20 Then
            Tank.Mode = 0
        End If
    End If
ElseIf HoverState = HOVER_HEIGHT + 25 Then 'Tank is returned to the ground
    Tank.Y = Tank.Y + 3
    If Tank.Y > FLOOR - Tank.Height Then
        Tank.Y = FLOOR - Tank.Height
        HoverState = 0
        Tank.Picture = 0
    End If
ElseIf HoverState >= 100 And HoverState <= 115 Then 'Double Hover Up
    Tank.Y = Tank.Y - 3
    HoverState = HoverState + 1
ElseIf HoverState > 115 Then 'Double Hover Down
    Tank.Y = Tank.Y + 3
    If Tank.Y > FLOOR - Tank.Height Then
        Tank.Y = FLOOR - Tank.Height
        HoverState = 0
        Tank.Picture = 0
    End If
End If
If BoostOn = True Then 'You are boosting
    If HoverState = 0 Then BoostOn = False 'You can't boost when you're not hovering
    intBoost = intBoost + 3
    If intBoost > 90 Then 'You have maxed out your boost
        BoostOn = False
        OverHeat = True 'You can't boost again until this is false
    End If
    Tank.X = Tank.X + 15
    If Tank.X > 550 Then 'You're past the border of where your tank can move
        Tank.X = 550
        PlySound ("buzzer")
    End If
    If TANK_TYPE = 2 Then
        Tank.Picture = 3
    End If
ElseIf intBoost > 0 Then 'Reduce the boost
    intBoost = intBoost - 1
    If intBoost < 30 Then 'You are no longer overheating
        OverHeat = False
    End If
End If

If HoverState = 0 And TANK_TYPE = 4 Then 'Water non-hovering
    If Tank.SpeedY = 0 Then Tank.SpeedY = 1
    If Tank.SpeedY > 0 Then
        Tank.Y = Tank.Y + 1
        If Tank.Y >= FLOOR - Tank.Height Then
            Tank.SpeedY = -1
        End If
    ElseIf Tank.SpeedY < 0 Then
        Tank.Y = Tank.Y - 1
        If Tank.Y <= 310 - Tank.Height Then
            Tank.SpeedY = 1
        End If
    End If
End If

If AllRangeMode = True Then 'All Range Mode is enabled, you can move up and down freely
    If bUp = True Then 'Move the tank up
        Tank.Y = Tank.Y - TANK_SPEED
        If Tank.Y < 0 Then 'You've gone beyond the border of the level
            Tank.Y = 0
            PlySound ("buzzer")
        End If
    ElseIf bDown = True Then 'Move the tank down
        If Tank.Y <= FLOOR - Tank.Height Then
            Tank.Y = Tank.Y + TANK_SPEED
        Else
            Tank.Y = FLOOR - Tank.Height 'You've gone beyond the border of the level
            PlySound ("buzzer")
        End If
    ElseIf Tank.Mode < 8 Then 'Oscilate the tank up and down
        Tank.Mode = Tank.Mode + 1
        Tank.Y = Tank.Y - 1
    ElseIf Tank.Mode <= 16 Then
        Tank.Mode = Tank.Mode + 1
        Tank.Y = Tank.Y + 1
        If Tank.Mode = 16 Then Tank.Mode = 0
    End If
    If AllRangeBoost > 0 Then 'Right after you've hit the all range mode, you are being boosted up
        If Tank.Y >= 250 Then
            Tank.Y = Tank.Y - 6
        Else
            AllRangeBoost = 0
        End If
    End If
End If

'Change the tread frame on the tank
If TANK_TYPE <> 3 And TANK_TYPE <> 4 And TANK_TYPE <> 6 And LevelFinished = False Then  'Not the truck
    If Tank.Picture = 0 Then
        Tank.Picture = 1
    ElseIf Tank.Picture = 1 Then
        Tank.Picture = 0
    End If
End If

If TankBlink > 0 Then 'Reduce the blinking of the tank
    TankBlink = TankBlink - 1
End If
'Draw the tank
'If the mask is not drawn, a translucent effect occurs indicating the tank is in cloak mode
If bCloak = False And (TankBlink = 0 Or TankBlink Mod 4 = 0) Then 'Not in "blink mode", draw the mask
    Select Case TANK_TYPE
        Case 0
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM(Tank.Picture).hdc, 0, 0, vbSrcAnd)
        Case 1
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM2(Tank.Picture).hdc, 0, 0, vbSrcAnd)
        Case 2
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM3(Tank.Picture).hdc, 0, 0, vbSrcAnd)
        Case 3
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM4(Tank.Picture).hdc, 0, 0, vbSrcAnd)
        Case 4
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM5(Tank.Picture).hdc, 0, 0, vbSrcAnd)
        Case 5
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM6(Tank.Picture).hdc, 0, 0, vbSrcAnd)
        Case 6
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM7(Tank.Picture).hdc, 0, 0, vbSrcAnd)
        Case 7
            Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTankM8(Tank.Picture).hdc, 0, 0, vbSrcAnd)
    End Select
End If
'Draw the sprite of the tank
Select Case TANK_TYPE
    Case 0
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank(Tank.Picture).hdc, 0, 0, vbSrcPaint)
    Case 1
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank2(Tank.Picture).hdc, 0, 0, vbSrcPaint)
    Case 2
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank3(Tank.Picture).hdc, 0, 0, vbSrcPaint)
    Case 3
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank4(Tank.Picture).hdc, 0, 0, vbSrcPaint)
    Case 4
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank5(Tank.Picture).hdc, 0, 0, vbSrcPaint)
    Case 5
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank6(Tank.Picture).hdc, 0, 0, vbSrcPaint)
    Case 6
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank7(Tank.Picture).hdc, 0, 0, vbSrcPaint)
    Case 7
        Call BitBlt(Me.hdc, Tank.X, Tank.Y, Tank.Width, Tank.Height, picTank8(Tank.Picture).hdc, 0, 0, vbSrcPaint)
End Select
End Sub
Private Sub DrawBullet()
'Draw and move the bullet
On Error Resume Next
For q = 1 To 10 'Go through each bullet
    If Bullet(q).Visible = True Then
        Bullet(q).X = Bullet(q).X + Bullet(q).Speed
        Bullet(q).Y = Bullet(q).Y - Bullet(q).SpeedY
        If Bullet(q).X >= Me.ScaleWidth Or Bullet(q).Y + Bullet(q).Height <= 0 Then 'Past the edge of the screen
            Bullet(q).Visible = False
        End If
        For i = 1 To 100 'Check bullet against all enemy sprites
            If Enemy(i).Visible = True Then
                'If bullet has collided with the enemy
                If Hit(Bullet(q), Enemy(i)) = True And Enemy(i).Picture <> 25 And Enemy(i).Invincible = False Then
                    Enemy(i).HP = Enemy(i).HP - 1
                    If Enemy(i).HP <= 0 Then 'Enemy is destroyed
                        Enemy(i).Visible = False
                    End If
                    Bullet(q).Visible = False 'Get rid of the bullet
                    If Enemy(i).Picture = 8 Then 'HP Power up
                        intHP = intHP + 20 'Boost health
                        PlySound ("bing")
                    ElseIf Enemy(i).Picture = 16 Then 'Strong ice wall, reduce to weak ice wall
                        Enemy(i).Picture = 17
                        Enemy(i).Width = picSprite(17).ScaleWidth
                        Enemy(i).Height = picSprite(17).ScaleHeight
                        Enemy(i).Y = FLOOR - Enemy(i).Height
                        PlySound ("explo2")
                        Enemy(i).Visible = True
                    ElseIf Enemy(i).Picture = 20 Then 'pillar
                        Enemy(i).Picture = 21
                        Enemy(i).Width = picSprite(21).ScaleWidth
                        Enemy(i).Height = picSprite(21).ScaleHeight
                        Enemy(i).Y = FLOOR - Enemy(i).Height
                        PlySound ("explo2")
                        Enemy(i).Visible = True
                    ElseIf Enemy(i).Picture = 19 Then 'glass
                        Call CreateExplosion(Enemy(i))
                        PlySound ("glassbreak")
                    ElseIf Enemy(i).Picture = 45 Then 'Turret
                        Call CreateExplosion(Enemy(i))
                        Enemy(i).Picture = 46
                        Enemy(i).Width = picSprite(46).ScaleWidth
                        Enemy(i).Height = picSprite(46).ScaleHeight
                        Enemy(i).Y = FLOOR - Enemy(i).Height
                        Enemy(i).Visible = True
                        PlySound ("explo2")
                    ElseIf Enemy(i).Picture = 39 Then 'Rapid fire
                        Enemy(i).Visible = False
                        RapidFire = True 'Turn on rapid fire
                        PlySound ("bing")
                    ElseIf Enemy(i).Picture = 28 Then 'All range mode
                        PlySound ("allrange")
                        AllRangeMode = True
                        If HoverState = 0 Then
                            AllRangeBoost = 30
                        End If
                        HoverState = 0
                        If TANK_TYPE <> 3 Then
                            Tank.Picture = 2 'Hover
                        End If
                        FLOOR = 320
                        Enemy(i).Visible = False
                    Else 'Regular
                        PlySound ("explo2")
                        Call CreateExplosion(Enemy(i))
                    End If
                    Exit For
                End If
            End If
        Next 'i
        'Draw the bullet
        If Bullet(q).Picture = 0 Then 'Horizontal bullet
            Call BitBlt(Me.hdc, Bullet(q).X, Bullet(q).Y, Bullet(q).Width, Bullet(q).Height, picBulletM.hdc, 0, 0, vbSrcAnd)
            Call BitBlt(Me.hdc, Bullet(q).X, Bullet(q).Y, Bullet(q).Width, Bullet(q).Height, picBullet.hdc, 0, 0, vbSrcPaint)
        Else 'Vertical bullet
            Call BitBlt(Me.hdc, Bullet(q).X, Bullet(q).Y, Bullet(q).Width, Bullet(q).Height, picBullet2M.hdc, 0, 0, vbSrcAnd)
            Call BitBlt(Me.hdc, Bullet(q).X, Bullet(q).Y, Bullet(q).Width, Bullet(q).Height, picBullet2.hdc, 0, 0, vbSrcPaint)
        End If
    End If
Next 'i
End Sub
Private Sub DrawExplo()
'Draw the explosions
On Error Resume Next
For i = 1 To 5
    If Explosion(i).Visible = True Then
        Explosion(i).Mode = Explosion(i).Mode - 1
        If Explosion(i).Mode = 0 Then 'Hide the explosion since it has been displayed for the maximum amount of time
            Explosion(i).Visible = False
        End If
        Call BitBlt(Me.hdc, Explosion(i).X, Explosion(i).Y, Explosion(i).Width, Explosion(i).Height, picExplosionM(Explosion(i).Picture).hdc, 0, 0, vbSrcAnd)
        Call BitBlt(Me.hdc, Explosion(i).X, Explosion(i).Y, Explosion(i).Width, Explosion(i).Height, picExplosion(Explosion(i).Picture).hdc, 0, 0, vbSrcPaint)
    End If
Next 'i

If intShake > 0 Then Call Shake 'Shake the form
End Sub
Private Sub PlaneBoss()
'Handle movement and drawing of the plane boss
On Error Resume Next
If Boss.Mode = 0 Then
    Boss.Visible = True
    Boss.Width = picBoss(4).ScaleWidth
    Boss.Height = picBoss(4).ScaleHeight
    Boss.Picture = 4
    Boss.HP = 8
    Boss.Mode = 1
    Boss.X = Me.ScaleWidth
    Boss.Y = 25
    Boss.Speed = 4
ElseIf Boss.Mode = 1 Then 'Bomb dropping mode
    Boss.X = Boss.X - Boss.Speed
    If Boss.X <= (500 + (Boss.Speed - 4) * 30) And Boss.X <= Me.ScaleWidth - picMissile(7).ScaleWidth - 2 Then
        Enemy(1).Visible = True
        Enemy(1).Picture = 32
        Enemy(1).SecondPic = 7
        Enemy(1).Width = picMissile(7).ScaleWidth
        Enemy(1).Height = picMissile(7).ScaleHeight
        Enemy(1).Y = Boss.Y + Boss.Height
        Enemy(1).X = Boss.X + 20
        Enemy(1).Speed = 0
        Enemy(1).SpeedY = 6
        Enemy(1).Mode = 5
        Boss.Mode = 2
    End If
ElseIf Boss.Mode = 2 Then 'Bomb dropping mode
    Boss.X = Boss.X - Boss.Speed
    If Boss.X <= 350 Then
        Enemy(2).Visible = True
        Enemy(2).Picture = 32
        Enemy(2).SecondPic = 7
        Enemy(2).Width = picMissile(7).ScaleWidth
        Enemy(2).Height = picMissile(7).ScaleHeight
        Enemy(2).Y = Boss.Y + Boss.Height
        Enemy(2).X = Boss.X + 20
        Enemy(2).Speed = 0
        If Boss.HP = 8 Then
            Enemy(2).SpeedY = 5
        Else
            Enemy(2).SpeedY = 3
        End If
        Enemy(2).Mode = 5
        Boss.Mode = 3
    End If
ElseIf Boss.Mode = 3 Then 'Bomb dropping mode end
    Boss.X = Boss.X - Boss.Speed
    If Boss.X <= 0 - Boss.Width Then
        Boss.Mode = 4
        Boss.Picture = 5
        Boss.Width = picBoss(5).ScaleWidth
        Boss.Height = picBoss(5).ScaleHeight
        Boss.X = Me.ScaleWidth
        Boss.Y = 310 - Boss.Height
        Boss.SpeedY = Boss.Speed - 2
    End If
ElseIf Boss.Mode = 4 Then
    Boss.X = Boss.X - Boss.Speed
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y + Boss.Height <= 0 Then
        For i = 3 To Boss.Speed + 1 'Make volcanic rocks
            Enemy(i).Visible = True
            Dim intRand As Long
            Randomize
            intRand = Int(Rnd * 2)
            Enemy(i).Picture = 30 + intRand
            Enemy(i).Width = picSprite(Enemy(i).Picture).ScaleWidth
            Enemy(i).Height = picSprite(Enemy(i).Picture).ScaleHeight
            Enemy(i).X = Me.ScaleWidth
            intRand = -200 + Int(Rnd * 350)
            Enemy(i).Y = intRand
            Enemy(i).Draw = True
            Enemy(i).Speed = 4
        Next 'i
        intShake = 200
        intAsh = 200
        PlySound ("volcano")
        Boss.Mode = 5
    End If
ElseIf Boss.Mode = 5 Then
    Dim bVisible As Boolean
    bVisible = True
    For i = 3 To Boss.Speed + 1
        If Enemy(i).Visible = True Then
            bVisible = False
            Exit For
        End If
    Next 'i
    If bVisible = True Then
        Boss.Mode = 6
        Boss.X = Me.ScaleWidth + 100
        Boss.Y = 15
        Boss.SpeedY = 8
        Boss.Width = picBoss(4).ScaleWidth
        Boss.Height = picBoss(4).ScaleHeight
        Boss.Picture = 4
    End If
ElseIf Boss.Mode = 6 Then 'Suicide mode
    If Boss.X + Boss.Width <= 0 Then
        Boss.Mode = 1
        Boss.X = Me.ScaleWidth
        Boss.Y = 25
    End If
    Boss.Y = Boss.Y + Boss.SpeedY
    Boss.X = Boss.X - Boss.Speed - 4
    If Boss.Y + Boss.Height >= 300 Then
        Boss.Mode = 7
    End If
ElseIf Boss.Mode = 7 Then 'Suicide mode
    Boss.Y = Boss.Y - Boss.SpeedY
    Boss.X = Boss.X - Boss.Speed - 4
    If Boss.Y <= 100 Then
        Boss.Mode = 6
    End If
ElseIf Boss.Mode = 8 Then 'Retreat mode
    Boss.X = Boss.X + Boss.Speed + 4
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 1
        Boss.Y = 25
    End If
ElseIf Boss.Mode = 1000 Then 'Win mode
    Boss.Visible = False
    LevelFinished = True
    intBlackBars = Me.ScaleHeight - FLOOR
    Boss.Mode = 1001
End If

For i = 1 To 10
    If Bullet(i).Visible = True And Boss.Visible = True Then
        If Hit(Bullet(i), Boss) = True Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Boss.Speed = Boss.Speed + 1
            Call CreateExplosion(Boss, True)
            PlySound ("explo2")
            If Boss.HP <= 0 Then
                Boss.Mode = 1000
            End If
            If Boss.Mode = 6 Or Boss.Mode = 7 Then
                Boss.Mode = 8
            End If
        End If
    End If
Next 'i

If Tank.Blink > 0 Then
    Tank.Blink = Tank.Blink - 1
End If

If Hit(Boss, Tank) = True And Tank.Blink = 0 And Boss.Visible = True Then
    Call CreateExplosion(Boss)
    Tank.Blink = 15
    PlySound ("explo2")
    intHP = intHP - 20
    If intHP < 0 Then
        Call GameOver
    End If
End If

If Boss.Visible = True Then
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBossM(Boss.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBoss(Boss.Picture).hdc, 0, 0, vbSrcPaint)
End If

End Sub
Private Sub DrawBuildings()
'Handle the drawing the rectangular building on the city levels
On Error Resume Next
Dim intNoDraw(0 To 100) As Long 'Regions of the city to not draw (aka gaps)
Dim maxNoDraw As Long 'How many of these regions exist

maxNoDraw = 0
For i = 1 To 100
    If Enemy(i).Visible = True Or (Enemy(i).Picture = 28 And Enemy(i).X <= 0) Then 'The sprite is on the screen
        If (Enemy(i).Picture >= 48 And Enemy(i).Picture <= 50) Or Enemy(i).Picture = 28 Then 'The sprite is a gap
            If Enemy(i).X <= Me.ScaleWidth Then
                If Enemy(i).Picture <> 28 Then
                    intNoDraw(maxNoDraw) = Enemy(i).X - 3
                    intNoDraw(maxNoDraw + 1) = Enemy(i).X + Enemy(i).Width - 6
                    maxNoDraw = maxNoDraw + 2
                Else
                    intNoDraw(maxNoDraw) = Enemy(i).X - 3
                    intNoDraw(maxNoDraw + 1) = Me.ScaleWidth + 100
                    maxNoDraw = maxNoDraw + 2
                End If
            End If
        End If
    End If
Next 'i
If maxNoDraw = 0 Then
    Line (0, 300)-Step(Me.ScaleWidth, 20), COLGRAY, BF
    Line (0, 300)-Step(Me.ScaleWidth, 2), COLBLACK, BF
Else
    Line (0, 300)-Step(intNoDraw(0), 20), COLGRAY, BF
    Line (0, 300)-Step(intNoDraw(0), 2), COLBLACK, BF
    If intNoDraw(maxNoDraw - 1) <= Me.ScaleWidth Then
        Line (intNoDraw(maxNoDraw - 1), 300)-Step(Me.ScaleWidth - intNoDraw(maxNoDraw - 1), 20), COLGRAY, BF
        Line (intNoDraw(maxNoDraw - 1), 300)-Step(Me.ScaleWidth - intNoDraw(maxNoDraw - 1), 2), COLBLACK, BF
    End If
    If maxNoDraw > 2 Then
        For i = 0 To maxNoDraw - 2
            If i Mod 2 = 1 Then
                Line (intNoDraw(i), 300)-Step((intNoDraw(i + 1)) - (intNoDraw(i)), 20), COLGRAY, BF
                Line ((intNoDraw(i)), 300)-Step((intNoDraw(i + 1)) - (intNoDraw(i)), 2), COLBLACK, BF
            End If
        Next 'i
    End If
End If
End Sub
Private Sub BoatBoss()
'Draw and move the boat boss
On Error Resume Next
If Boss.Mode = 0 Then
    Boss.Picture = 6
    Boss.Width = picBoss(6).ScaleWidth
    Boss.Height = picBoss(6).ScaleHeight
    Boss.Y = FLOOR - Boss.Height
    Boss.X = Me.ScaleWidth
    Boss.HP = 10
    Boss.Mode = 1
    Boss.Visible = True
    Boss.Speed = 2
ElseIf Boss.Mode = 1 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.X <= Me.ScaleWidth - Boss.Width Then
        Boss.Mode = 2
    End If
ElseIf Boss.Mode = 2 Then
    Enemy(11).Visible = True
    Enemy(11).Draw = True
    Enemy(11).Picture = 51
    Enemy(11).Width = picSprite(51).ScaleWidth
    Enemy(11).Height = picSprite(51).ScaleHeight
    Enemy(11).HP = 1
    Enemy(11).Invincible = True
    Enemy(11).Damage = 10
    Enemy(11).Mode = 0
    Enemy(11).X = Boss.X - Enemy(11).Width
    Enemy(11).Y = Boss.Y + 79
    Boss.Mode = 3
ElseIf Boss.Mode <= 15 Then
    Boss.Mode = Boss.Mode + 1
ElseIf Boss.Mode = 16 Then
    Boss.X = Boss.X + Boss.Speed
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 17
        Boss.Width = picBoss(7).ScaleWidth
        Boss.Height = picBoss(7).ScaleHeight
        Boss.Picture = 7
    End If
ElseIf Boss.Mode <= 25 Then
    Boss.Mode = Boss.Mode + 1
ElseIf Boss.Mode = 26 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.X + Boss.Width <= Me.ScaleWidth Then
        Boss.Mode = 27
    End If
ElseIf Boss.Mode = 27 Then
    Enemy(11).Visible = True
    Enemy(11).Draw = True
    Enemy(11).Picture = 52
    Enemy(11).Width = picSprite(52).ScaleWidth
    Enemy(11).Height = picSprite(52).ScaleHeight
    Enemy(11).HP = 1
    Enemy(11).Invincible = True
    Enemy(11).Damage = 10
    Enemy(11).Mode = 0
    Enemy(11).X = Boss.X - Enemy(11).Width
    Enemy(11).Y = Boss.Y + 77
    Boss.Mode = 28
ElseIf Boss.Mode <= 35 Then
    Boss.Mode = Boss.Mode + 1
ElseIf Boss.Mode = 36 Then
    Boss.X = Boss.X + Boss.Speed
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 37
        Boss.Width = picBoss(8).ScaleWidth
        Boss.Height = picBoss(8).ScaleHeight
        Boss.Picture = 8
    End If
ElseIf Boss.Mode <= 48 Then
    Boss.Mode = Boss.Mode + 1
ElseIf Boss.Mode = 49 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.X + Boss.Width <= Me.ScaleWidth Then
        Enemy(11).Visible = True
        Enemy(11).Draw = True
        Enemy(11).Picture = 53
        Enemy(11).Width = picSprite(53).ScaleWidth
        Enemy(11).Height = picSprite(53).ScaleHeight
        Enemy(11).HP = 1
        Enemy(11).Invincible = True
        Enemy(11).Damage = 10
        Enemy(11).Mode = 0
        Enemy(11).X = Boss.X - Enemy(11).Width
        Enemy(11).Y = Boss.Y + 75
        Boss.Mode = 50
    End If
ElseIf Boss.Mode <= 60 Then
    Boss.Mode = Boss.Mode + 1
ElseIf Boss.Mode = 61 Then
    Boss.X = Boss.X + Boss.Speed
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 62
        Boss.Width = picBoss(6).ScaleWidth
        Boss.Height = picBoss(6).ScaleHeight
        Boss.Picture = 6
    End If
ElseIf Boss.Mode <= 73 Then
    Boss.Mode = Boss.Mode + 1
ElseIf Boss.Mode = 74 Then
    Boss.Mode = 1
End If

If Enemy(10).Visible = False Then
    Enemy(10).Visible = True
    If Enemy(10).Picture = 54 Then
        Enemy(10).Picture = 55
        Enemy(10).Width = picSprite(55).ScaleWidth
        Enemy(10).Height = picSprite(55).ScaleHeight
        Enemy(10).X = 0 - Enemy(10).Width + 10
    Else
        Enemy(10).Picture = 54
        Enemy(10).X = Me.ScaleWidth
        Enemy(10).Width = picSprite(55).ScaleWidth
        Enemy(10).Height = picSprite(55).ScaleHeight
    End If
    Randomize
    Enemy(10).Y = -30 + Int(Rnd * 100)
    Enemy(10).Draw = True
    Enemy(10).HP = 1
    Enemy(10).Mode = 0
    Enemy(10).Damage = 10
    Enemy(10).Speed = 4
    Enemy(10).SpeedY = 0
End If

If Boss.Speed > 8 Then Boss.Speed = 8

'x1=60 x2=140 y1=0 y2 = 65

For i = 1 To 10
    If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 142 And Bullet(i).X + Bullet(i).Width >= Boss.X + 60 And Bullet(i).Y <= Boss.Y + 65 And Bullet(i).Y + Bullet(i).Height >= Boss.Y Then
        Boss.HP = Boss.HP - 1
        PlySound ("explo2")
        Bullet(i).Visible = False
        Call CreateExplosion(Bullet(i), True)
        Boss.Speed = Boss.Speed + 1
        If Boss.HP <= 0 Then
            PlySound ("explosion")
            Boss.Visible = False
            LevelFinished = True
            intBlackBars = Me.ScaleHeight - FLOOR
            Boss.Mode = 1001
        End If
    End If
Next 'i

If Hit(Tank, Boss) = True And TankBlink = 0 Then
    intHP = intHP - 10
    If intHP < 0 Then
        Call GameOver
    End If
    Call CreateExplosion(Tank, True)
    TankBlink = 30
    PlySound ("explo2")
    intShake = 15
End If

If Boss.Visible = True Then
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBossM(Boss.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBoss(Boss.Picture).hdc, 0, 0, vbSrcPaint)
End If

End Sub
Private Sub BaseBoss()
'Draw and move the underwater base boss
On Error Resume Next
Randomize
If Boss.Mode = 0 Then
    Boss.Picture = 9
    Boss.Width = picBoss(9).ScaleWidth
    Boss.Height = picBoss(9).ScaleHeight
    Boss.X = Me.ScaleWidth
    Boss.Y = FLOOR - Boss.Height
    Boss.Speed = 6
    Boss.HP = 8
    Boss.Visible = True
    BossT.Picture = 19
    BossT.Visible = False
    BossT.Width = picBoss(19).ScaleWidth
    BossT.Height = picBoss(19).ScaleHeight
    BossT.HP = 1
    Boss.Mode = 1
ElseIf Boss.Mode = 1 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.X + Boss.Width <= Me.ScaleWidth Then
        Boss.Mode = 4
        Boss.Picture = 10
        Boss.Width = picBoss(10).ScaleWidth
        Boss.Height = picBoss(10).ScaleHeight
    End If
ElseIf Boss.Mode = 2 Then
    Boss.X = Boss.X + Boss.Speed
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 1
    End If
ElseIf Boss.Mode >= 4 Then
    Boss.Mode = Boss.Mode + 1
    If Boss.Mode Mod (38 - Boss.Speed) = 0 Then
        Dim intNewSprite As Long
        intNewSprite = FindNewSprite
        PlySound ("explo2")
        With Enemy(intNewSprite)
            .Visible = True
            Dim intRand As Long
            intRand = Int(Rnd * 4)
            intRand = 51 + intRand
            If intRand = 54 Then
                intRand = 52
                .Mode = 3
            Else
                .Mode = 0
            End If
            .Picture = intRand
            .Width = picSprite(intRand).ScaleWidth
            .Height = picSprite(intRand).ScaleHeight
            .Invincible = True
            .SpeedY = 0
            .Speed = 6
            Select Case .Picture
                Case 51
                    .X = Boss.X + 10
                    .Y = Boss.Y + 154
                Case 52
                    .Y = 182
                    .X = Boss.X - .Width
                Case 53
                    .Y = Boss.Y + 12
                    .X = Boss.X + 62
            End Select
        End With
    End If
    If Boss.Mode Mod 75 = 0 And BossT.Visible = False Then
        BossT.Visible = True
        If BossT.Picture = 19 Then
            BossT.Picture = 20
            BossT.Height = picBoss(20).ScaleHeight
            BossT.Width = picBoss(20).ScaleWidth
            BossT.Y = FLOOR
            BossT.Mode = 0
            BossT.X = Int(Rnd * 175) + 225
        Else
            BossT.Picture = 19
            BossT.Height = picBoss(19).ScaleHeight
            BossT.Width = picBoss(19).ScaleWidth
            BossT.Y = 0 - BossT.Height
            BossT.Mode = 2
            BossT.X = Int(Rnd * 175) + 225
        End If
    End If
End If

If BossT.Visible = True Then
    If BossT.Mode = 0 Then
        BossT.Y = BossT.Y - (Boss.Speed / 2) + 2
        If BossT.Y + BossT.Height <= FLOOR Then
            BossT.Mode = 1
            intNewSprite = FindNewSprite
            With Enemy(intNewSprite)
                .Visible = True
                .Picture = 47
                .Width = picSprite(47).ScaleWidth
                .Height = picSprite(47).ScaleHeight
                .Invincible = True
                .Mode = 0
                .Speed = 6
                .SpeedY = 0
                .Y = BossT.Y + 5
                .X = BossT.X - .Width
                .Draw = True
                .Damage = 10
            End With
        End If
    ElseIf BossT.Mode = 1 Then
        BossT.Y = BossT.Y + (Boss.Speed / 2) - 2
        If BossT.Y >= FLOOR Then
            BossT.Visible = False
        End If
    ElseIf BossT.Mode = 2 Then
        BossT.Y = BossT.Y + (Boss.Speed / 2) - 2
        If BossT.Y >= 0 Then
            BossT.Mode = 3
            intNewSprite = FindNewSprite
            With Enemy(intNewSprite)
                .Visible = True
                .Picture = 47
                .Width = picSprite(47).ScaleWidth
                .Height = picSprite(47).ScaleHeight
                .Invincible = True
                .Mode = 0
                .Speed = 6
                .SpeedY = 0
                .Y = BossT.Y + BossT.Height - 5
                .X = BossT.X - .Width
                .Draw = True
                .Damage = 10
            End With
        End If
    ElseIf BossT.Mode = 3 Then
        BossT.Y = BossT.Y - (Boss.Speed / 2) + 2
        If BossT.Y + BossT.Height <= 0 Then
            BossT.Visible = False
        End If
    End If
    For i = 1 To 10
        If Bullet(i).Visible = True Then
            If Hit(Bullet(i), BossT) = True Then
                Bullet(i).Visible = False
                BossT.Visible = False
                Call CreateExplosion(Bullet(i))
                Call CreateExplosion(Boss, True)
                PlySound ("explo2")
                Boss.HP = Boss.HP - 1
                Boss.Speed = Boss.Speed + 1
                Boss.Picture = 9
                Boss.Width = picBoss(9).ScaleWidth
                Boss.Height = picBoss(9).ScaleHeight
                Boss.Mode = 2
                If Boss.HP = 0 Then
                    Boss.Mode = 3
                    Boss.Visible = False
                    PlySound ("explosion")
                    LevelFinished = True
                    intBlackBars = Me.ScaleHeight - FLOOR
                End If
            End If
        End If
    Next 'i
    Call BitBlt(Me.hdc, BossT.X, BossT.Y, BossT.Width, BossT.Height, picBossM(BossT.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, BossT.X, BossT.Y, BossT.Width, BossT.Height, picBoss(BossT.Picture).hdc, 0, 0, vbSrcPaint)
End If

If Hit(Tank, Boss) = True And TankBlink = 0 Then
    intHP = intHP - 10
    If intHP < 0 Then
        Call GameOver
    End If
    Call CreateExplosion(Tank, True)
    TankBlink = 30
    PlySound ("explo2")
    intShake = 15
End If

If Hit(Tank, BossT) = True And TankBlink = 0 Then
    intHP = intHP - 10
    If intHP < 0 Then
        Call GameOver
    End If
    Call CreateExplosion(Tank, True)
    TankBlink = 30
    PlySound ("explo2")
    intShake = 15
End If

If Boss.Visible = True Then
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBossM(Boss.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBoss(Boss.Picture).hdc, 0, 0, vbSrcPaint)
End If

End Sub
Private Sub FinalBoss()
'Draw and move the final mecha boss
On Error Resume Next
Randomize
If Boss.Mode = 0 Then
    Boss.Picture = 11
    Boss.Width = picBoss(11).ScaleWidth
    Boss.Height = picBoss(11).ScaleHeight
    Boss.X = Me.ScaleWidth
    Boss.Y = 100
    Boss.HP = 3
    Boss.Mode = 1
    Boss.Speed = 5
    Boss.SpeedY = 3
    Boss.Visible = True
    Boss.SecondPic = 0
ElseIf Boss.Mode = 1 Then
    Boss.X = Boss.X - Boss.Speed
    If Boss.X + Boss.Width <= Me.ScaleWidth Then
        Boss.Mode = 2
    End If
ElseIf Boss.Mode = 2 Then 'Fire lasers
    If Enemy(10).Visible = False Then
        With Enemy(10)
            .Visible = True
            .Picture = 47
            .Width = picSprite(47).ScaleWidth
            .Height = picSprite(47).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 15
            .Mode = 0
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y + Boss.SpeedY
    If Boss.Y >= FLOOR - 25 Then
        Boss.Mode = 3
    End If
ElseIf Boss.Mode = 3 Then
    If Enemy(10).Visible = False Then
        With Enemy(10)
            .Visible = True
            .Picture = 47
            .Width = picSprite(47).ScaleWidth
            .Height = picSprite(47).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 15
            .Mode = 0
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y <= -100 Then
        Boss.Mode = 2
    End If
ElseIf Boss.Mode = 4 Then 'Fire missiles
    If Enemy(10).Visible = False Then
        With Enemy(10)
            .Visible = True
            .Picture = 52
            .Width = picSprite(52).ScaleWidth
            .Height = picSprite(52).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 165
            .Mode = 2 'Curve down
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y + Boss.SpeedY
    If Boss.Y >= FLOOR - 25 Then
        Boss.Mode = 5
    End If
ElseIf Boss.Mode = 5 Then
    If Enemy(11).Visible = False Then
        With Enemy(11)
            .Visible = True
            .Picture = 52
            .Width = picSprite(52).ScaleWidth
            .Height = picSprite(52).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 155
            .Mode = 3 ' Curve up
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y <= -100 Then
        Boss.Mode = 4
    End If
ElseIf Boss.Mode = 6 Then 'Fire lasers
    If Enemy(10).Visible = False Then
        With Enemy(10)
            .Visible = True
            .Picture = 47
            .Width = picSprite(47).ScaleWidth
            .Height = picSprite(47).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 186
            .Mode = 2 'Curve down
            .Damage = 10
        End With
    End If
    If Boss.Y >= 50 And Enemy(13).Visible = False Then
        With Enemy(13)
            .Visible = True
            .Picture = 47
            .Width = picSprite(47).ScaleWidth
            .Height = picSprite(47).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 186
            .Mode = 2 'Curve down
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y + Boss.SpeedY
    If Boss.Y >= FLOOR - 25 Then
        Boss.Mode = 7
    End If
ElseIf Boss.Mode = 7 Then
    If Enemy(11).Visible = False Then
        With Enemy(11)
            .Visible = True
            .Picture = 47
            .Width = picSprite(47).ScaleWidth
            .Height = picSprite(47).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 186
            .Mode = 3 ' Curve up
            .Damage = 10
        End With
    End If
    If Boss.Y <= 100 And Enemy(12).Visible = False Then
        With Enemy(12)
            .Visible = True
            .Picture = 47
            .Width = picSprite(47).ScaleWidth
            .Height = picSprite(47).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 186
            .Mode = 3 ' Curve up
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y <= -100 Then
        Boss.Mode = 6
    End If
ElseIf Boss.Mode = 8 Then 'Fire plasma balls
    If Enemy(51).Visible = False And Boss.Y + 260 <= FLOOR Then
        With Enemy(51)
            .Visible = True
            .Picture = 43
            .Width = picSprite(43).ScaleWidth
            .Height = picSprite(43).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = 5
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 245
            .Mode = Int(Rnd * 75)
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y + Boss.SpeedY
    If Boss.Y >= FLOOR Then
        Boss.Mode = 9
    End If
ElseIf Boss.Mode = 9 Then
    If Enemy(50).Visible = False And Boss.Y + 260 <= FLOOR Then
        With Enemy(50)
            .Visible = True
            .Picture = 43
            .Width = picSprite(43).ScaleWidth
            .Height = picSprite(43).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = 5
            .SpeedY = 0
            .X = Boss.X - .Width
            .Y = Boss.Y + 245
            .Mode = Int(Rnd * 75)
            .Damage = 10
        End With
    End If
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y <= -265 Then
        Boss.Mode = 8
    End If
ElseIf Boss.Mode = 10 Then 'Fire speedy missiles
    If Boss.SecondPic > 0 Then
        Boss.SecondPic = Boss.SecondPic - 1
        Boss.Picture = 18
    Else
        Boss.Picture = 15
    End If
    If Enemy(50).Visible = False Then
        With Enemy(50)
            .Visible = True
            .Picture = 61
            .Width = picSprite(61).ScaleWidth
            .Height = picSprite(61).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed - 5
            .SpeedY = 0
            .X = Boss.X + 106 - .Width
            .Y = Boss.Y + 75
            .Mode = 0
            .Damage = 10
        End With
        Boss.SecondPic = 10
    End If
    If Boss.Y >= 50 And Enemy(53).Visible = False Then
        With Enemy(53)
            .Visible = True
            .Picture = 61
            .Width = picSprite(61).ScaleWidth
            .Height = picSprite(61).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed - 5
            .SpeedY = 0
            .X = Boss.X + 106 - .Width
            .Y = Boss.Y + 75
            .Mode = 0
            .Damage = 10
        End With
        Boss.SecondPic = 10
    End If
    Boss.Y = Boss.Y + Boss.SpeedY
    If Boss.Y >= FLOOR - 25 Then
        Boss.Mode = 11
    End If
ElseIf Boss.Mode = 11 Then
    If Boss.SecondPic > 0 Then
        Boss.SecondPic = Boss.SecondPic - 1
        Boss.Picture = 18
    Else
        Boss.Picture = 15
    End If
    If Enemy(51).Visible = False Then
        With Enemy(51)
            .Visible = True
            .Picture = 61
            .Width = picSprite(61).ScaleWidth
            .Height = picSprite(61).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X + 106 - .Width
            .Y = Boss.Y + 75
            .Mode = 0 ' Curve up
            .Damage = 10
        End With
        Boss.SecondPic = 10
    End If
    If Boss.Y + 95 <= FLOOR And Enemy(52).Visible = False Then
        With Enemy(52)
            .Visible = True
            .Picture = 61
            .Width = picSprite(61).ScaleWidth
            .Height = picSprite(61).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed
            .SpeedY = 0
            .X = Boss.X + 106 - .Width
            .Y = Boss.Y + 75
            .Mode = 0 ' Curve up
            .Damage = 10
        End With
        Boss.SecondPic = 10
    End If
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y <= -100 Then
        Boss.Mode = 10
    End If
ElseIf Boss.Mode = 12 Then 'Fire up from back cannon
    Boss.Y = Boss.Y + Boss.SpeedY
    If Boss.Y >= FLOOR - 25 Then
        Boss.Mode = 13
    End If
ElseIf Boss.Mode = 13 Then
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y <= -100 Then
        Boss.Mode = 12
    End If
    If Enemy(51).Visible = False Then
        With Enemy(51)
            .Visible = True
            .Picture = 62
            .Width = picSprite(62).ScaleWidth
            .Height = picSprite(62).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = 0
            .SpeedY = -14
            .SecondPic = (14 - Boss.HP)
            .X = Boss.X + 250
            .Y = Boss.Y + 200 - .Height
            .Mode = 0
            .Damage = 10
        End With
    End If
ElseIf Boss.Mode = 14 Then 'Fire giant ball of fire
    Boss.Y = Boss.Y + Boss.SpeedY
    If Boss.Y >= FLOOR - 25 Then
        Boss.Mode = 15
    End If
    If Enemy(51).Visible = False Then
        With Enemy(51)
            .Visible = True
            .Picture = 64
            .Width = picSprite(64).ScaleWidth
            .Height = picSprite(64).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed - 15
            .SpeedY = 0
            .X = Boss.X + 215 - .Width
            .Y = Boss.Y + 50
            intRand = Int(Rnd * 3)
            .Mode = (intRand * 2)
            .Damage = 10
        End With
    End If
ElseIf Boss.Mode = 15 Then
    Boss.Y = Boss.Y - Boss.SpeedY
    If Boss.Y <= -100 Then
        Boss.Mode = 14
    End If
    If Enemy(51).Visible = False And Boss.Y + 50 <= FLOOR - 70 Then
        With Enemy(51)
            .Visible = True
            .Picture = 64
            .Width = picSprite(64).ScaleWidth
            .Height = picSprite(64).ScaleHeight
            .HP = 1
            .Invincible = True
            .Draw = True
            .Speed = Boss.Speed - 20
            .SpeedY = 0
            .X = Boss.X + 215 - .Width
            .Y = Boss.Y + 50
            intRand = Int(Rnd * 3)
            .Mode = (intRand * 2)
            .Damage = 10
        End With
    End If
End If


If Boss.Mode = 2 Or Boss.Mode = 3 Then
    'x1 = 8, x2 = 75, y1=4, y2 = 34
    For i = 1 To 10
        If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 75 And Bullet(i).X + Bullet(i).Width >= Boss.X + 8 And Bullet(i).Y <= Boss.Y + 34 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 4 Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Call CreateExplosion(Bullet(i))
            PlySound ("explo2")
            Boss.Speed = Boss.Speed + 1
            Boss.SpeedY = Boss.SpeedY + 1
            If Boss.HP = 0 Then
                Boss.SpeedY = Boss.SpeedY - 2
                For q = 2 To 5
                    Explosion(q).Visible = True
                    Explosion(q).X = Boss.X + (75 * (q - 1))
                    Explosion(q).Y = Boss.Y + (75 * (q - 1))
                    Explosion(q).Picture = 2
                    Explosion(q).Mode = 10
                    Explosion(q).Width = picExplosion(2).ScaleWidth
                    Explosion(q).Height = picExplosion(2).ScaleHeight
                Next 'q
                PlySound ("explosion")
                Boss.Mode = 4
                Boss.HP = 4
                Boss.Picture = Boss.Picture + 1
            End If
        End If
    Next 'i
ElseIf Boss.Mode = 4 Or Boss.Mode = 5 Then
    'x1 = 10, x2 = 50, y1=150, y2 = 175
    For i = 1 To 10
        If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 50 And Bullet(i).X + Bullet(i).Width >= Boss.X + 10 And Bullet(i).Y <= Boss.Y + 175 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 150 Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Call CreateExplosion(Bullet(i))
            PlySound ("explo2")
            Boss.Speed = Boss.Speed + 1
            Boss.SpeedY = Boss.SpeedY + 1
            If Boss.HP = 0 Then
                Boss.SpeedY = Boss.SpeedY - 2
                Boss.Mode = 6
                Boss.HP = 5
                Boss.Picture = Boss.Picture + 1
                PlySound ("explosion")
                For q = 2 To 5
                    Explosion(q).Visible = True
                    Explosion(q).X = Boss.X + (75 * (q - 1))
                    Explosion(q).Y = Boss.Y + (75 * (q - 1))
                    Explosion(q).Picture = 2
                    Explosion(q).Mode = 10
                    Explosion(q).Width = picExplosion(2).ScaleWidth
                    Explosion(q).Height = picExplosion(2).ScaleHeight
                Next 'q
            End If
        End If
    Next 'i
ElseIf Boss.Mode = 6 Or Boss.Mode = 7 Then
    '180-205 y, 4-50 x
    For i = 1 To 10
        If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 50 And Bullet(i).X + Bullet(i).Width >= Boss.X + 4 And Bullet(i).Y <= Boss.Y + 205 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 180 Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Call CreateExplosion(Bullet(i))
            PlySound ("explo2")
            Boss.Speed = Boss.Speed + 1
            Boss.SpeedY = Boss.SpeedY + 1
            If Boss.HP = 0 Then
                Boss.SpeedY = Boss.SpeedY - 2
                Boss.Mode = 8
                Boss.HP = 6
                PlySound ("explosion")
                Boss.Picture = Boss.Picture + 1
                For q = 2 To 5
                    Explosion(q).Visible = True
                    Explosion(q).X = Boss.X + (75 * (q - 1))
                    Explosion(q).Y = Boss.Y + (75 * (q - 1))
                    Explosion(q).Picture = 2
                    Explosion(q).Mode = 10
                    Explosion(q).Width = picExplosion(2).ScaleWidth
                    Explosion(q).Height = picExplosion(2).ScaleHeight
                Next 'q
            End If
        End If
    Next 'i
ElseIf Boss.Mode = 8 Or Boss.Mode = 9 Then
    '235-265 y, 6-45 x
    For i = 1 To 10
        If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 45 And Bullet(i).X + Bullet(i).Width >= Boss.X + 6 And Bullet(i).Y <= Boss.Y + 265 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 235 Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Call CreateExplosion(Bullet(i))
            PlySound ("explo2")
            Boss.Speed = Boss.Speed + 1
            Boss.SpeedY = Boss.SpeedY + 1
            If Boss.HP = 0 Then
                PlySound ("explosion")
                Boss.SpeedY = Boss.SpeedY - 2
                Boss.Mode = 10
                Boss.HP = 6
                Boss.Picture = Boss.Picture + 1
                For q = 2 To 5
                    Explosion(q).Visible = True
                    Explosion(q).X = Boss.X + (75 * (q - 1))
                    Explosion(q).Y = Boss.Y + (75 * (q - 1))
                    Explosion(q).Picture = 2
                    Explosion(q).Mode = 10
                    Explosion(q).Width = picExplosion(2).ScaleWidth
                    Explosion(q).Height = picExplosion(2).ScaleHeight
                Next 'q
            End If
        End If
    Next 'i
ElseIf Boss.Mode = 10 Or Boss.Mode = 11 Then
    '47-91 y, 107-188 x
    For i = 1 To 10
        If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 188 And Bullet(i).X + Bullet(i).Width >= Boss.X + 107 And Bullet(i).Y <= Boss.Y + 91 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 47 Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Call CreateExplosion(Bullet(i))
            PlySound ("explo2")
            Boss.Speed = Boss.Speed + 1
            Boss.SpeedY = Boss.SpeedY + 1
            If Boss.HP = 0 Then
                PlySound ("glassbreak")
                Boss.SpeedY = Boss.SpeedY - 3
                Boss.Mode = 12
                Boss.HP = 7
                Boss.Picture = 16
                For q = 2 To 5
                    Explosion(q).Visible = True
                    Explosion(q).X = Boss.X + (75 * (q - 1))
                    Explosion(q).Y = Boss.Y + (75 * (q - 1))
                    Explosion(q).Picture = 2
                    Explosion(q).Mode = 10
                    Explosion(q).Width = picExplosion(2).ScaleWidth
                    Explosion(q).Height = picExplosion(2).ScaleHeight
                Next 'q
            End If
        End If
    Next 'i
ElseIf Boss.Mode = 12 Or Boss.Mode = 13 Then
    '200-280 y, 235-282 x
    For i = 1 To 10
        If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 282 And Bullet(i).X + Bullet(i).Width >= Boss.X + 235 And Bullet(i).Y <= Boss.Y + 280 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 200 Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Call CreateExplosion(Bullet(i))
            PlySound ("explo2")
            Boss.Speed = Boss.Speed + 1
            PlySound ("explosion")
            Boss.SpeedY = Boss.SpeedY + 1
            If Boss.HP = 0 Then
                PlySound ("explosion")
                Boss.SpeedY = Boss.SpeedY - 3
                Boss.Mode = 14
                Boss.HP = 10
                Boss.Picture = Boss.Picture + 1
                For q = 2 To 5
                    Explosion(q).Visible = True
                    Explosion(q).X = Boss.X + (75 * (q - 1))
                    Explosion(q).Y = Boss.Y + (75 * (q - 1))
                    Explosion(q).Picture = 2
                    Explosion(q).Mode = 10
                    Explosion(q).Width = picExplosion(2).ScaleWidth
                    Explosion(q).Height = picExplosion(2).ScaleHeight
                Next 'q
            End If
        End If
    Next 'i
ElseIf Boss.Mode = 14 Or Boss.Mode = 15 Then
    '96-266 y, 96-236 x
    For i = 1 To 10
        If Bullet(i).Visible = True And Bullet(i).X <= Boss.X + 236 And Bullet(i).X + Bullet(i).Width >= Boss.X + 96 And Bullet(i).Y <= Boss.Y + 266 And Bullet(i).Y + Bullet(i).Height >= Boss.Y + 96 Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Call CreateExplosion(Bullet(i))
            PlySound ("explo2")
            Boss.Speed = Boss.Speed + 1
            Boss.SpeedY = Boss.SpeedY + 1
            If Boss.HP = 0 Then
                PlySound ("explosion")
                Boss.Mode = 16
                Boss.Visible = False
                LevelFinished = True
                intBlackBars = Me.ScaleHeight - FLOOR
                For q = 2 To 5
                    Explosion(q).Visible = True
                    Explosion(q).X = Boss.X + (75 * (q - 1))
                    Explosion(q).Y = Boss.Y + (75 * (q - 1))
                    Explosion(q).Picture = 2
                    Explosion(q).Mode = 10
                    Explosion(q).Width = picExplosion(2).ScaleWidth
                    Explosion(q).Height = picExplosion(2).ScaleHeight
                Next 'q
            End If
        End If
    Next 'i
End If

If Boss.Visible = True Then
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBossM(Boss.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBoss(Boss.Picture).hdc, 0, 0, vbSrcPaint)
End If

End Sub
Private Sub DrawForeground()
'Draw the HUD on the bottom of the screen such as the health meter and the cannon fire meter

On Error Resume Next
'Draw ground
Line (0, 320)-Step(Me.ScaleWidth, Me.ScaleHeight - 320), GROUND_COLOR, BF
'Health box
Call BitBlt(Me.hdc, 5, 325, picHealth.ScaleWidth, picHealth.ScaleHeight, picHealthM.hdc, 0, 0, vbSrcAnd)
Call BitBlt(Me.hdc, 5, 325, picHealth.ScaleWidth, picHealth.ScaleHeight, picHealth.hdc, 0, 0, vbSrcPaint)
'Handle cannon ready state
If CannonState < 3 Then
    CannonWait = CannonWait + 1
    If CannonWait >= BULLET_COOLDOWN Or (RapidFire = True And CannonWait >= 4) Then
        CannonState = CannonState + 1
        CannonWait = 0
    End If
End If
'Cannon ready draw
Call BitBlt(Me.hdc, Me.ScaleWidth - picCannon(0).ScaleWidth - 15, 325, picCannon(0).ScaleWidth, picCannon(0).ScaleHeight, picCannonM(CannonState).hdc, 0, 0, vbSrcAnd)
Call BitBlt(Me.hdc, Me.ScaleWidth - picCannon(0).ScaleWidth - 15, 325, picCannon(0).ScaleWidth, picCannon(0).ScaleHeight, picCannon(CannonState).hdc, 0, 0, vbSrcPaint)
'Splat logo draw
Call BitBlt(Me.hdc, Me.ScaleWidth / 2 - 125, 320, picSplat.ScaleWidth, picSplat.ScaleHeight, picSplatM.hdc, 0, 0, vbSrcAnd)
Call BitBlt(Me.hdc, Me.ScaleWidth / 2 - 125, 320, picSplat.ScaleWidth, picSplat.ScaleHeight, picSplat.hdc, 0, 0, vbSrcPaint)
'Draw All Range Mode text
If AllRangeMode = True Then
    Dim strText As String
    Me.Font = "Arial"
    Me.FontBold = True
    Me.FontSize = 16
    Me.ForeColor = COLRED
    strText = "ALL RANGE MODE"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 75, 320 + picSplat.ScaleHeight - 40, strText, Len(strText))
End If
'Draw Boost
If bBoost = True Then
    Call BitBlt(Me.hdc, Me.ScaleWidth / 2 - 115 + picSplat.ScaleWidth, 320, picBoost.ScaleWidth, picBoost.ScaleHeight, picBoostM.hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Me.ScaleWidth / 2 - 115 + picSplat.ScaleWidth, 320, picBoost.ScaleWidth, picBoost.ScaleHeight, picBoost.hdc, 0, 0, vbSrcPaint)
    Dim tempColor As Long
    If OverHeat = False Then
        tempColor = COLGREEN
    Else
        tempColor = COLRED
    End If
    Line (Me.ScaleWidth / 2 - 112 + picSplat.ScaleWidth, 316 + picBoost.Height - intBoost)-Step(picBoost.ScaleWidth - 6, intBoost), tempColor, BF
    Me.Font = "Arial"
    Me.FontBold = True
    Me.FontSize = 12
    Me.ForeColor = COLBLACK
    strText = "B"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 112 + picSplat.ScaleWidth + 4, 324, strText, Len(strText))
    strText = "O"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 112 + picSplat.ScaleWidth + 4, 342, strText, Len(strText))
    strText = "O"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 112 + picSplat.ScaleWidth + 4, 360, strText, Len(strText))
    strText = "S"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 112 + picSplat.ScaleWidth + 4, 378, strText, Len(strText))
    strText = "T"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 112 + picSplat.ScaleWidth + 4, 396, strText, Len(strText))
End If

'Draw cloak bar
If AllowCloak = True Then
    Call BitBlt(Me.hdc, Me.ScaleWidth / 2 - 130, 320, picBoost.ScaleWidth, picBoost.ScaleHeight, picBoostM.hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Me.ScaleWidth / 2 - 130, 320, picBoost.ScaleWidth, picBoost.ScaleHeight, picBoost.hdc, 0, 0, vbSrcPaint)
    If intCloak < 76 Then
        tempColor = COLRED
    Else
        tempColor = COLGREEN
    End If
    Line (Me.ScaleWidth / 2 - 128, 316 + picBoost.Height - (intCloak))-Step(picBoost.ScaleWidth - 6, (intCloak)), tempColor, BF
    Me.Font = "Arial"
    Me.FontBold = True
    Me.FontSize = 12
    Me.ForeColor = COLBLACK
    strText = "C"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 122, 324, strText, Len(strText))
    strText = "L"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 122, 342, strText, Len(strText))
    strText = "O"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 122, 360, strText, Len(strText))
    strText = "A"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 122, 378, strText, Len(strText))
    strText = "K"
    Call TextOut(Me.hdc, Me.ScaleWidth / 2 - 122, 396, strText, Len(strText))
End If

'Draw HP
strText = CStr(intHP)
Me.Font = "Arial"
Me.FontBold = True
Me.FontSize = 12
Me.ForeColor = COLGREEN
Call TextOut(Me.hdc, 55, 360, strText, Len(strText))

End Sub
Private Sub RedAlert()
'If you have 0 HP, flash the screen red every so often
On Error Resume Next
intRedAlert = intRedAlert + 1
If intRedAlert Mod 2 = 0 Then
intShake = 2
    Else
intShake = 1
End If
If intRedAlert > 29 Then
    intRedAlert = 0
    Line (0, 0)-Step(Me.ScaleWidth, Me.ScaleHeight), COLRED, BF
End If
End Sub
Private Sub doDemo()
'Automatically moves the tank in demo mode
On Error Resume Next
Me.Font = "Arial Black"
Me.ForeColor = COLRED
Me.FontSize = 26
Me.FontBold = True
Dim strText As String
strText = "Demo Mode - Press Enter"
Call TextOut(Me.hdc, 125, 150, strText, Len(strText))

intDemo = intDemo + 1

If NextLevel = "level2.ini" Then 'Playing the first level
    Select Case intDemo
        Case 75
            MakeBullet
        Case 100
            bUp = True
            HoverState = 1
        Case 105
            bRight = True
        Case 155
            bRight = False
        Case 180
            MakeBullet
        Case 200
            bLeft = True
        Case 240
            bLeft = False
        Case 250
            MakeBullet
        Case 252
            HoverState = 1
        Case 375
            HoverState = 1
        Case 385
            bRight = True
        Case 420
            bRight = False
        Case 440
            MakeBullet
        Case 490
            MakeBullet
        Case 540
            bLeft = True
        Case 585
            MakeBullet
        Case 590
            bLeft = False
        Case 685
            HoverState = 1
            bRight = True
        Case 760
            bRight = False
        Case 780
            bLeft = True
        Case 840
            bLeft = False
            HoverState = 1
        Case 841
            bRight = True
        Case 870
            bRight = False
        Case 914
            HoverState = 1
        Case 1000 'End of demo mode
            frmIntro.timeDemo.Enabled = True
            Unload Me
            frmIntro.Show
            bRunning = False
            Exit Sub
    End Select
ElseIf NextLevel = "level4.ini" Then 'Playing the snow level
    If intDemo < 130 Then intDemo = 130
    Select Case intDemo
        Case 192
            bRight = True
        Case 224
            MakeBullet
            bRight = False
        Case 262
            bLeft = True
        Case 269
            MakeBullet
        Case 285
            bLeft = False
        Case 314
            MakeBullet
        Case 340
            bLeft = True
        Case 355
            bLeft = False
        Case 369
            HoverState = 1
        Case 378
            bRight = True
            MakeBullet
        Case 416
            MakeBullet
        Case 424
            bRight = False
        Case 444
            bLeft = True
        Case 473
            MakeBullet
        Case 484
            bLeft = False
        Case 494
            bRight = True
        Case 505
            bRight = False
        Case 510
            bLeft = True
        Case 524
            bLeft = False
        Case 545
            bRight = True
        Case 555
            MakeBullet
        Case 565
            bRight = False
        Case 566
            bLeft = True
        Case 573
            bLeft = False
        Case 574
            bRight = True
        Case 589
            HoverState = 1
        Case 600
            bRight = False
        Case 604
            MakeBullet
        Case 613
            bLeft = True
        Case 621
            bDown = True
        Case 622
            bDown = False
        Case 635
            bLeft = False
        Case 648
            bLeft = True
        Case 651
            bLeft = False
        Case 654
            HoverState = 1
        Case 669
            MakeBullet
        Case 674
            bLeft = True
        Case 675
            bDown = True
        Case 676
            bDown = False
        Case 692
            bLeft = False
        Case 722
            bLeft = True
        Case 724
            MakeBullet
        Case 735
            bLeft = False
        Case 755
            bRight = True
        Case 787
            bRight = False
        Case 788
            MakeBullet
        Case 816
            bLeft = True
        Case 848
            bLeft = False
        Case 854
            HoverState = 1
        Case 865
            bRight = True
        Case 909
            bRight = False
        Case 931
            bRight = True
        Case 932
            MakeBullet
        Case 933
            bRight = False
        Case 951
            bLeft = True
        Case 964
            bLeft = False
        Case 972
            MakeBullet
        Case 1002
            bLeft = True
        Case 1005
            bLeft = False
        Case 1010
            MakeBullet
        Case 1030
            bLeft = True
        Case 1035
            bLeft = False
        Case 1145
            bRight = True
        Case 1161
            bRight = False
        Case 1172
            bLeft = True
        Case 1186
            bLeft = False
        Case 1215
            frmIntro.timeDemo.Enabled = True
            bRunning = False
            Unload Me
            frmIntro.Show
            Exit Sub
    End Select
End If
End Sub
Private Sub MarsBoss()
'Draw and move the tank for the mars boss
On Error Resume Next
If Boss.Mode = 0 Then
    Boss.Visible = True
    Boss.Picture = 21
    Boss.Width = picBoss(21).ScaleWidth
    Boss.Height = picBoss(21).ScaleHeight
    Boss.Y = FLOOR - Boss.Height
    Boss.X = Me.ScaleWidth
    Boss.HP = 15
    Boss.Speed = 5
    Boss.Mode = 1
ElseIf Boss.Mode = 1 Then 'Drive out
    Boss.X = Boss.X - Boss.Speed
    If Boss.X + Boss.Width <= Me.ScaleWidth - 25 Then
        Boss.Mode = 10
    End If
ElseIf Boss.Mode = 2 Then
    Boss.X = Boss.X + Boss.Speed
    If Boss.X >= Me.ScaleWidth Then
        Boss.Mode = 1
    End If
ElseIf Boss.Mode >= 10 Then
    If (Boss.Mode Mod 100 <= 50) Then
        Boss.X = Boss.X - Boss.Speed
    Else
        Boss.X = Boss.X + Boss.Speed
    End If
    If Boss.Mode Mod 50 = 0 Then
        Dim intNewSprite As Long
        intNewSprite = FindNewSprite
        With Enemy(intNewSprite)
            .Visible = True
            .Picture = 47
            .Invincible = True
            .Width = picSprite(47).ScaleWidth
            .Height = picSprite(47).ScaleHeight
            .X = Boss.X - .Width
            .Y = Boss.Y + 16
            .Draw = True
            .HP = 1
            .Damage = 10
            .Speed = Boss.Speed
            .Mode = 0
            .SpeedY = 0
        End With
        PlySound ("explo2")
    End If
    If Boss.Mode Mod 90 = 0 And Boss.SecondPic = 0 Then
        Boss.SecondPic = 1
        PlySound ("hover")
    End If
    If Boss.Mode Mod 100 = 0 Then
        Dim intRand As Long
        Randomize
        intRand = Int(Rnd * 1)
        With Enemy(intNewSprite)
            .Visible = True
            .Picture = 54 + intRand
            .Invincible = False
            .SpeedY = 0
            .Width = picSprite(.Picture).ScaleWidth
            .Height = picSprite(.Picture).ScaleHeight
            .Mode = 0
            If .Picture = 54 Then
                .X = Me.ScaleWidth
                .Speed = 6
            Else
                .X = 0
                .Speed = -6
            End If
            .Y = 15
            .Draw = True
            .HP = 1
            .Damage = 10
        End With
    End If
    If Boss.Mode Mod 150 = 0 Then
        intRand = Int(Rnd * 40)
        intNewSprite = FindNewSprite
        With Enemy(intNewSprite)
            .Visible = True
            .Picture = 43
            .Invincible = False
            .SpeedY = 0
            .Width = picSprite(.Picture).ScaleWidth
            .Height = picSprite(.Picture).ScaleHeight
            .Mode = intRand
            .Speed = 6
            .X = Me.ScaleWidth
            .Y = 210
            .Draw = True
            .HP = 1
            .Damage = 10
        End With
    End If
    Boss.Mode = Boss.Mode + 1
End If

If Boss.SecondPic = 0 Then 'Not hovering
    Boss.Picture = Boss.Picture + 1
    If Boss.Picture >= 23 Then
        Boss.Picture = 21
    End If
Else 'Hover
    Boss.Picture = 23
    If Boss.SecondPic > 0 And Boss.SecondPic < 20 Then
        Boss.SecondPic = Boss.SecondPic + 1
        Boss.Y = Boss.Y - Boss.Speed
    ElseIf Boss.SecondPic < 25 Then
        Boss.SecondPic = Boss.SecondPic + 1
    ElseIf Boss.SecondPic < 45 Then
        Boss.Y = Boss.Y + Boss.Speed
        Boss.SecondPic = Boss.SecondPic + 1
    Else
        Boss.Y = FLOOR - Boss.Height
        Boss.SecondPic = 0
    End If
End If

For i = 1 To 10
    If Bullet(i).Visible = True Then
        If Hit(Bullet(i), Boss) = True Then
            Bullet(i).Visible = False
            Boss.HP = Boss.HP - 1
            Boss.Speed = Boss.Speed
            Call CreateExplosion(Boss, True)
            PlySound ("explosion")
            If Boss.HP Mod 4 = 0 Then
                Boss.Speed = Boss.Speed + 1
                Boss.Mode = 2
            End If
            If Boss.HP = 0 Then
                Boss.Visible = False
                Boss.Mode = 9
                LevelFinished = True
                intBlackBars = Me.ScaleHeight - FLOOR
            End If
        End If
    End If
Next 'i

If Hit(Tank, Boss) = True And TankBlink = 0 Then
    intHP = intHP - 10
    If intHP < 0 Then
        Call GameOver
    End If
    Call CreateExplosion(Tank, True)
    TankBlink = 30
    PlySound ("explo2")
    intShake = 15
End If

If Boss.Visible = True Then
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBossM(Boss.Picture).hdc, 0, 0, vbSrcAnd)
    Call BitBlt(Me.hdc, Boss.X, Boss.Y, Boss.Width, Boss.Height, picBoss(Boss.Picture).hdc, 0, 0, vbSrcPaint)
End If
End Sub




