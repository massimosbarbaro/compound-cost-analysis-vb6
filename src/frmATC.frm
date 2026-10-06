VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCosti 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Costi Mescola"
   ClientHeight    =   10305
   ClientLeft      =   870
   ClientTop       =   405
   ClientWidth     =   12720
   Icon            =   "frmATC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   12720
   Begin VB.TextBox txtSP 
      Height          =   285
      Left            =   5040
      TabIndex        =   369
      Text            =   "Text1"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtMPD 
      Height          =   285
      Left            =   6240
      TabIndex        =   368
      Text            =   "Text1"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtMPPS 
      Height          =   285
      Left            =   8880
      TabIndex        =   367
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtMPS 
      Height          =   285
      Left            =   8040
      TabIndex        =   366
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtMPC 
      Height          =   285
      Left            =   7200
      TabIndex        =   365
      Text            =   "Text1"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtMP 
      Height          =   285
      Left            =   5640
      TabIndex        =   364
      Text            =   "Text1"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtRCC5 
      Height          =   285
      Left            =   11520
      TabIndex        =   363
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtRCC3 
      Height          =   285
      Left            =   8280
      TabIndex        =   362
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtRCC2 
      Height          =   285
      Left            =   5040
      TabIndex        =   361
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtRCE10 
      Height          =   285
      Left            =   8280
      TabIndex        =   360
      Text            =   "e/kg"
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txtRCM01 
      Height          =   285
      Left            =   8280
      TabIndex        =   359
      Text            =   "e/kg"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtRCA10 
      DataMember      =   "e/kg"
      Height          =   285
      Left            =   5160
      TabIndex        =   358
      Text            =   "e/kg"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtRCB11 
      Height          =   285
      Left            =   5280
      TabIndex        =   357
      Text            =   "e/kg"
      Top             =   7560
      Width           =   375
   End
   Begin VB.TextBox txtRCQ01 
      Height          =   285
      Left            =   1680
      TabIndex        =   356
      Text            =   "e/kg"
      Top             =   8400
      Width           =   375
   End
   Begin VB.TextBox txtRCBJ 
      Height          =   285
      Left            =   1680
      TabIndex        =   355
      Text            =   "e/kg"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox txtRCB00 
      Height          =   285
      Left            =   1680
      TabIndex        =   354
      Text            =   "e/kg"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtRCS01 
      Height          =   285
      Left            =   1680
      TabIndex        =   353
      Text            =   "e/kg"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtCCMa514 
      Height          =   285
      Left            =   10560
      TabIndex        =   349
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtCStMa514 
      Height          =   285
      Left            =   10560
      TabIndex        =   348
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtRCC1 
      Height          =   285
      Left            =   1680
      TabIndex        =   347
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   10
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   346
      Text            =   ""
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtCSB11 
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   339
      Text            =   "12,25"
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox txtCSSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   338
      Text            =   "12,25"
      Top             =   6840
      Width           =   495
   End
   Begin VB.TextBox txtTSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   337
      Text            =   "12,25"
      Top             =   6915
      Width           =   495
   End
   Begin VB.TextBox txtCRSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   336
      Text            =   "12,25"
      Top             =   6915
      Width           =   495
   End
   Begin VB.TextBox txtCRCB11 
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   335
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtTCB11 
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   334
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtCSB11 
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   333
      Text            =   "12,25"
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox txtCSSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   332
      Text            =   "12,25"
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox txtTSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   331
      Text            =   "12,25"
      Top             =   6195
      Width           =   495
   End
   Begin VB.TextBox txtCRSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   330
      Text            =   "12,25"
      Top             =   6195
      Width           =   495
   End
   Begin VB.TextBox txtCRCB11 
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   329
      Text            =   "12,25"
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox txtTCB11 
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   328
      Text            =   "12,25"
      Top             =   6600
      Width           =   495
   End
   Begin VB.TextBox txtCSB11 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   327
      Text            =   "12,25"
      Top             =   5760
      Width           =   495
   End
   Begin VB.TextBox txtCSSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   326
      Text            =   "12,25"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtTSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   325
      Text            =   "12,25"
      Top             =   5475
      Width           =   495
   End
   Begin VB.TextBox txtCRSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   324
      Text            =   "12,25"
      Top             =   5475
      Width           =   495
   End
   Begin VB.TextBox txtCRCB11 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   323
      Text            =   "12,25"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtTCB11 
      Height          =   285
      Index           =   1
      Left            =   5400
      TabIndex        =   322
      Text            =   "12,25"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtCSQ01 
      Height          =   285
      Index           =   3
      Left            =   8400
      TabIndex        =   321
      Text            =   "12,25"
      Top             =   9405
      Width           =   495
   End
   Begin VB.TextBox txtCSSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   8400
      TabIndex        =   320
      Text            =   "12,25"
      Top             =   9000
      Width           =   495
   End
   Begin VB.TextBox txtTSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   10080
      TabIndex        =   319
      Text            =   "12,25"
      Top             =   9000
      Width           =   495
   End
   Begin VB.TextBox txtCRSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   9480
      TabIndex        =   318
      Text            =   "12,25"
      Top             =   9000
      Width           =   495
   End
   Begin VB.TextBox txtCRCQ01 
      Height          =   285
      Index           =   3
      Left            =   9480
      TabIndex        =   317
      Text            =   "12,25"
      Top             =   9405
      Width           =   495
   End
   Begin VB.TextBox txtTCQ01 
      Height          =   285
      Index           =   3
      Left            =   10080
      TabIndex        =   316
      Text            =   "12,25"
      Top             =   9405
      Width           =   495
   End
   Begin VB.TextBox txtCSQ01 
      Height          =   285
      Index           =   2
      Left            =   6000
      TabIndex        =   315
      Text            =   "12,25"
      Top             =   9165
      Width           =   495
   End
   Begin VB.TextBox txtCSSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   6000
      TabIndex        =   314
      Text            =   "12,25"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox txtTSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   313
      Text            =   "12,25"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox txtCRSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   7080
      TabIndex        =   312
      Text            =   "12,25"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox txtCRCQ01 
      Height          =   285
      Index           =   2
      Left            =   7080
      TabIndex        =   311
      Text            =   "12,25"
      Top             =   9165
      Width           =   495
   End
   Begin VB.TextBox txtTCQ01 
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   310
      Text            =   "12,25"
      Top             =   9165
      Width           =   495
   End
   Begin VB.TextBox txtCSQ01 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   309
      Text            =   "12,25"
      Top             =   9165
      Width           =   495
   End
   Begin VB.TextBox txtCSSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   308
      Text            =   "12,25"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox txtTSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   307
      Text            =   "12,25"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox txtCRSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   306
      Text            =   "12,25"
      Top             =   8760
      Width           =   495
   End
   Begin VB.TextBox txtCRCQ01 
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   305
      Text            =   "12,25"
      Top             =   9165
      Width           =   495
   End
   Begin VB.TextBox txtTCQ01 
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   304
      Text            =   "12,25"
      Top             =   9165
      Width           =   495
   End
   Begin VB.TextBox txtCSBJ 
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   303
      Text            =   "12,25"
      Top             =   1845
      Width           =   495
   End
   Begin VB.TextBox txtCSSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   2760
      TabIndex        =   302
      Text            =   "12,25"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtTSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   301
      Text            =   "12,25"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtCRSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   300
      Text            =   "12,25"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtCRCBJ 
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   299
      Text            =   "12,25"
      Top             =   1845
      Width           =   495
   End
   Begin VB.TextBox txtTCBJ 
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   298
      Text            =   "12,25"
      Top             =   1845
      Width           =   495
   End
   Begin VB.TextBox txtCSBJ 
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   297
      Text            =   "12,25"
      Top             =   1125
      Width           =   495
   End
   Begin VB.TextBox txtCSSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   296
      Text            =   "12,25"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtTSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   295
      Text            =   "12,25"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCRSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   294
      Text            =   "12,25"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txtCRCBJ 
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   293
      Text            =   "12,25"
      Top             =   1125
      Width           =   495
   End
   Begin VB.TextBox txtTCBJ 
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   292
      Text            =   "12,25"
      Top             =   1125
      Width           =   495
   End
   Begin VB.TextBox txtCSBJ 
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   291
      Text            =   "12,25"
      Top             =   405
      Width           =   495
   End
   Begin VB.TextBox txtCSSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   290
      Text            =   "12,25"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtTSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   289
      Text            =   "12,25"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtCRSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   288
      Text            =   "12,25"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtCRCBJ 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   287
      Text            =   "12,25"
      Top             =   405
      Width           =   495
   End
   Begin VB.TextBox txtTCBJ 
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   286
      Text            =   "12,25"
      Top             =   405
      Width           =   495
   End
   Begin VB.TextBox txtCSB00 
      Height          =   285
      Index           =   3
      Left            =   7080
      TabIndex        =   285
      Text            =   "12,25"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtCSSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   7080
      TabIndex        =   284
      Text            =   "12,25"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtTSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   8280
      TabIndex        =   283
      Text            =   "12,25"
      Top             =   195
      Width           =   495
   End
   Begin VB.TextBox txtCRSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   7680
      TabIndex        =   282
      Text            =   "12,25"
      Top             =   195
      Width           =   495
   End
   Begin VB.TextBox txtCRCB00 
      Height          =   285
      Index           =   3
      Left            =   7680
      TabIndex        =   281
      Text            =   "12,25"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtTCB00 
      Height          =   285
      Index           =   3
      Left            =   8280
      TabIndex        =   280
      Text            =   "12,25"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtCSB00 
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   279
      Text            =   "12,25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtCSSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   5040
      TabIndex        =   278
      Text            =   "12,25"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtTSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   6120
      TabIndex        =   277
      Text            =   "12,25"
      Top             =   -45
      Width           =   495
   End
   Begin VB.TextBox txtCRSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   276
      Text            =   "12,25"
      Top             =   -45
      Width           =   495
   End
   Begin VB.TextBox txtCRCB00 
      Height          =   285
      Index           =   2
      Left            =   5520
      TabIndex        =   275
      Text            =   "12,25"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtTCB00 
      Height          =   285
      Index           =   2
      Left            =   6120
      TabIndex        =   274
      Text            =   "12,25"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtCSB00 
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   273
      Text            =   "12,25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtCSSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   6600
      TabIndex        =   272
      Text            =   "12,25"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtTSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   7800
      TabIndex        =   271
      Text            =   "12,25"
      Top             =   75
      Width           =   495
   End
   Begin VB.TextBox txtCRSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   7200
      TabIndex        =   270
      Text            =   "12,25"
      Top             =   75
      Width           =   495
   End
   Begin VB.TextBox txtCRCB00 
      Height          =   285
      Index           =   1
      Left            =   7200
      TabIndex        =   269
      Text            =   "12,25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtTCB00 
      Height          =   285
      Index           =   1
      Left            =   7800
      TabIndex        =   268
      Text            =   "12,25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtCSA10 
      Height          =   285
      Index           =   3
      Left            =   11280
      TabIndex        =   267
      Text            =   "12,25"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtCSSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   11280
      TabIndex        =   266
      Text            =   "12,25"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtTCA10 
      Height          =   285
      Index           =   3
      Left            =   12360
      TabIndex        =   265
      Text            =   "12,25"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtCRCA10 
      Height          =   285
      Index           =   3
      Left            =   11760
      TabIndex        =   264
      Text            =   "12,25"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtCRSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   11760
      TabIndex        =   263
      Text            =   "12,25"
      Top             =   675
      Width           =   495
   End
   Begin VB.TextBox txtTSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   12360
      TabIndex        =   262
      Text            =   "12,25"
      Top             =   675
      Width           =   495
   End
   Begin VB.TextBox txtCSA10 
      Height          =   285
      Index           =   2
      Left            =   8640
      TabIndex        =   261
      Text            =   "12,25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtCSSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   8640
      TabIndex        =   260
      Text            =   "12,25"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtTCA10 
      Height          =   285
      Index           =   2
      Left            =   9720
      TabIndex        =   259
      Text            =   "12,25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtCRCA10 
      Height          =   285
      Index           =   2
      Left            =   9120
      TabIndex        =   258
      Text            =   "12,25"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtCRSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   9120
      TabIndex        =   257
      Text            =   "12,25"
      Top             =   75
      Width           =   495
   End
   Begin VB.TextBox txtTSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   9720
      TabIndex        =   256
      Text            =   "12,25"
      Top             =   75
      Width           =   495
   End
   Begin VB.TextBox txtCSA10 
      Height          =   285
      Index           =   1
      Left            =   11280
      TabIndex        =   255
      Text            =   "12,25"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtCSSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   11280
      TabIndex        =   254
      Text            =   "12,25"
      Top             =   -120
      Width           =   495
   End
   Begin VB.TextBox txtTCA10 
      Height          =   285
      Index           =   1
      Left            =   12360
      TabIndex        =   253
      Text            =   "12,25"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtCRCA10 
      Height          =   285
      Index           =   1
      Left            =   11760
      TabIndex        =   252
      Text            =   "12,25"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtCRSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   11760
      TabIndex        =   251
      Text            =   "12,25"
      Top             =   -45
      Width           =   495
   End
   Begin VB.TextBox txtTSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   12360
      TabIndex        =   250
      Text            =   "12,25"
      Top             =   -45
      Width           =   495
   End
   Begin VB.TextBox txtCSM01 
      Height          =   285
      Index           =   3
      Left            =   9960
      TabIndex        =   249
      Text            =   "12,25"
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox txtCSSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   9960
      TabIndex        =   248
      Text            =   "12,25"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtTCM01 
      Height          =   285
      Index           =   3
      Left            =   11640
      TabIndex        =   247
      Text            =   "12,25"
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox txtCRCM01 
      Height          =   285
      Index           =   3
      Left            =   11040
      TabIndex        =   246
      Text            =   "12,25"
      Top             =   6120
      Width           =   495
   End
   Begin VB.TextBox txtCRSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   11040
      TabIndex        =   245
      Text            =   "12,25"
      Top             =   5715
      Width           =   495
   End
   Begin VB.TextBox txtTSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   11640
      TabIndex        =   244
      Text            =   "12,25"
      Top             =   5715
      Width           =   495
   End
   Begin VB.TextBox txtCSM01 
      Height          =   285
      Index           =   2
      Left            =   9960
      TabIndex        =   243
      Text            =   "12,25"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox txtCSSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   9960
      TabIndex        =   242
      Text            =   "12,25"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtTCM01 
      Height          =   285
      Index           =   2
      Left            =   11640
      TabIndex        =   241
      Text            =   "12,25"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox txtCRCM01 
      Height          =   285
      Index           =   2
      Left            =   11040
      TabIndex        =   240
      Text            =   "12,25"
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox txtCRSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   11040
      TabIndex        =   239
      Text            =   "12,25"
      Top             =   4875
      Width           =   495
   End
   Begin VB.TextBox txtTSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   11640
      TabIndex        =   238
      Text            =   "12,25"
      Top             =   4875
      Width           =   495
   End
   Begin VB.TextBox txtCSM01 
      Height          =   285
      Index           =   1
      Left            =   9960
      TabIndex        =   237
      Text            =   "12,25"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCSSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   9960
      TabIndex        =   236
      Text            =   "12,25"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtTCM01 
      Height          =   285
      Index           =   1
      Left            =   11640
      TabIndex        =   235
      Text            =   "12,25"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCRCM01 
      Height          =   285
      Index           =   1
      Left            =   11040
      TabIndex        =   234
      Text            =   "12,25"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtCRSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   11040
      TabIndex        =   233
      Text            =   "12,25"
      Top             =   3915
      Width           =   495
   End
   Begin VB.TextBox txtTSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   11640
      TabIndex        =   232
      Text            =   "12,25"
      Top             =   3915
      Width           =   495
   End
   Begin VB.TextBox txtCSS01 
      Height          =   285
      Index           =   3
      Left            =   8880
      TabIndex        =   231
      Text            =   "12,25"
      Top             =   8565
      Width           =   495
   End
   Begin VB.TextBox txtCSSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   8880
      TabIndex        =   230
      Text            =   "12,25"
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox txtTCS01 
      Height          =   285
      Index           =   3
      Left            =   10560
      TabIndex        =   229
      Text            =   "12,25"
      Top             =   8565
      Width           =   495
   End
   Begin VB.TextBox txtCRCS01 
      Height          =   285
      Index           =   3
      Left            =   9960
      TabIndex        =   228
      Text            =   "12,25"
      Top             =   8565
      Width           =   495
   End
   Begin VB.TextBox txtCRSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   9960
      TabIndex        =   227
      Text            =   "12,25"
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox txtTSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   10560
      TabIndex        =   226
      Text            =   "12,25"
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox txtCSS01 
      Height          =   285
      Index           =   2
      Left            =   8880
      TabIndex        =   225
      Text            =   "12,25"
      Top             =   7725
      Width           =   495
   End
   Begin VB.TextBox txtCSSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   8880
      TabIndex        =   224
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtTCS01 
      Height          =   285
      Index           =   2
      Left            =   10560
      TabIndex        =   223
      Text            =   "12,25"
      Top             =   7725
      Width           =   495
   End
   Begin VB.TextBox txtCRCS01 
      Height          =   285
      Index           =   2
      Left            =   9960
      TabIndex        =   222
      Text            =   "12,25"
      Top             =   7725
      Width           =   495
   End
   Begin VB.TextBox txtCRSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   9960
      TabIndex        =   221
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtTSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   10560
      TabIndex        =   220
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtCSS01 
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   219
      Text            =   "12,25"
      Top             =   7725
      Width           =   495
   End
   Begin VB.TextBox txtCSSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   218
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtTCS01 
      Height          =   285
      Index           =   1
      Left            =   8160
      TabIndex        =   217
      Text            =   "12,25"
      Top             =   7725
      Width           =   495
   End
   Begin VB.TextBox txtCRCS01 
      Height          =   285
      Index           =   1
      Left            =   7560
      TabIndex        =   216
      Text            =   "12,25"
      Top             =   7725
      Width           =   495
   End
   Begin VB.TextBox txtCRSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   7560
      TabIndex        =   215
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtTSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   8160
      TabIndex        =   214
      Text            =   "12,25"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox txtTCB11 
      Height          =   285
      Index           =   0
      Left            =   5400
      TabIndex        =   210
      Text            =   "12,25"
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox txtRB11 
      Height          =   285
      Left            =   4920
      TabIndex        =   209
      Text            =   "e/kg"
      Top             =   7560
      Width           =   375
   End
   Begin VB.TextBox txtCRCB11 
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   208
      Text            =   "12,25"
      Top             =   8280
      Width           =   495
   End
   Begin VB.TextBox txtCRSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   207
      Text            =   "12,25"
      Top             =   7875
      Width           =   495
   End
   Begin VB.TextBox txtTSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   5400
      TabIndex        =   206
      Text            =   "12,25"
      Top             =   7875
      Width           =   495
   End
   Begin VB.TextBox txtTCBJ 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   203
      Text            =   "12,25"
      Top             =   7680
      Width           =   495
   End
   Begin VB.TextBox txtRBJ 
      Height          =   285
      Left            =   1320
      TabIndex        =   202
      Text            =   "e/kg"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox txtCRCBJ 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   201
      Text            =   "12,25"
      Top             =   7680
      Width           =   495
   End
   Begin VB.TextBox txtCRSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   200
      Text            =   "12,25"
      Top             =   7275
      Width           =   495
   End
   Begin VB.TextBox txtTSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   199
      Text            =   "12,25"
      Top             =   7275
      Width           =   495
   End
   Begin VB.TextBox txtTCQ01 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   198
      Text            =   "12,25"
      Top             =   9120
      Width           =   495
   End
   Begin VB.TextBox txtRQ01 
      Height          =   285
      Left            =   1320
      TabIndex        =   197
      Text            =   "e/kg"
      Top             =   8400
      Width           =   375
   End
   Begin VB.TextBox txtCRCQ01 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   196
      Text            =   "12,25"
      Top             =   9120
      Width           =   495
   End
   Begin VB.TextBox txtCRSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   195
      Text            =   "12,25"
      Top             =   8715
      Width           =   495
   End
   Begin VB.TextBox txtTSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   194
      Text            =   "12,25"
      Top             =   8715
      Width           =   495
   End
   Begin VB.TextBox txtTCB00 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   191
      Text            =   "12,25"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox txtRB00 
      Height          =   285
      Left            =   1320
      TabIndex        =   190
      Text            =   "e/kg"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtCRCB00 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   189
      Text            =   "12,25"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox txtCRSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   188
      Text            =   "12,25"
      Top             =   5835
      Width           =   495
   End
   Begin VB.TextBox txtTSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   187
      Text            =   "12,25"
      Top             =   5835
      Width           =   495
   End
   Begin VB.TextBox txtTSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   174
      Text            =   "12,25"
      Top             =   4515
      Width           =   495
   End
   Begin VB.TextBox txtCRSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   173
      Text            =   "12,25"
      Top             =   4515
      Width           =   495
   End
   Begin VB.TextBox txtCRCM01 
      Height          =   285
      Index           =   0
      Left            =   7680
      TabIndex        =   172
      Text            =   "12,25"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtRM01 
      Height          =   285
      Left            =   7800
      TabIndex        =   171
      Text            =   "e/kg"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtTCM01 
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   170
      Text            =   "12,25"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtTSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   169
      Text            =   "12,25"
      Top             =   4515
      Width           =   495
   End
   Begin VB.TextBox txtCRSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   168
      Text            =   "12,25"
      Top             =   4515
      Width           =   495
   End
   Begin VB.TextBox txtCRCA10 
      Height          =   285
      Index           =   0
      Left            =   4560
      TabIndex        =   167
      Text            =   "12,25"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtRA10 
      DataMember      =   "e/kg"
      Height          =   285
      Left            =   4680
      TabIndex        =   166
      Text            =   "e/kg"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtTCA10 
      Height          =   285
      Index           =   0
      Left            =   5160
      TabIndex        =   165
      Text            =   "12,25"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtTSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   162
      Text            =   "12,25"
      Top             =   4395
      Width           =   495
   End
   Begin VB.TextBox txtCRSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   161
      Text            =   "12,25"
      Top             =   4395
      Width           =   495
   End
   Begin VB.TextBox txtCRCS01 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   160
      Text            =   "12,25"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtRS01 
      Height          =   285
      Left            =   1320
      TabIndex        =   159
      Text            =   "e/kg"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtTCS01 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   158
      Text            =   "12,25"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtTCE10 
      Height          =   285
      Left            =   8280
      TabIndex        =   157
      Text            =   "12,25"
      Top             =   6720
      Width           =   495
   End
   Begin VB.TextBox txtRE10 
      Height          =   285
      Left            =   7800
      TabIndex        =   156
      Text            =   "e/kg"
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txtCRCE10 
      Height          =   285
      Left            =   7680
      TabIndex        =   155
      Text            =   "12,25"
      Top             =   6720
      Width           =   495
   End
   Begin VB.TextBox txtCRSE10 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   7680
      TabIndex        =   154
      Text            =   "12,25"
      Top             =   6315
      Width           =   495
   End
   Begin VB.TextBox txtTSE10 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   8280
      TabIndex        =   153
      Text            =   "12,25"
      Top             =   6315
      Width           =   495
   End
   Begin VB.TextBox txtTSC5 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   11880
      TabIndex        =   150
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCRSC5 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   11280
      TabIndex        =   149
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCRCC5 
      Height          =   285
      Left            =   11280
      TabIndex        =   148
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtRC5 
      Height          =   285
      Left            =   11160
      TabIndex        =   147
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtTCC5 
      Height          =   285
      Left            =   11880
      TabIndex        =   146
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtTSC3 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   8640
      TabIndex        =   143
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCRSC3 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   8040
      TabIndex        =   142
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCRCC3 
      Height          =   285
      Left            =   8040
      TabIndex        =   141
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtRC3 
      Height          =   285
      Left            =   7920
      TabIndex        =   140
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtTCC3 
      Height          =   285
      Left            =   8640
      TabIndex        =   139
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtTSC2 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   5400
      TabIndex        =   132
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCRSC2 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   4800
      TabIndex        =   131
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCRCC2 
      Height          =   285
      Left            =   4800
      TabIndex        =   130
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtRC2 
      Height          =   285
      Left            =   4680
      TabIndex        =   129
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtTCC2 
      Height          =   285
      Left            =   5400
      TabIndex        =   128
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtTCC1 
      Height          =   285
      Left            =   2160
      TabIndex        =   127
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtRC1 
      Height          =   285
      Left            =   1320
      TabIndex        =   123
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtCRCC1 
      Height          =   285
      Left            =   1560
      TabIndex        =   122
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtCRSC1 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   1560
      TabIndex        =   121
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtTSC1 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   2160
      TabIndex        =   120
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCSSC2 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   3720
      TabIndex        =   119
      Text            =   "0.68"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCSSC3 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   6960
      TabIndex        =   118
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCSSS01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   117
      Text            =   "12,25"
      Top             =   4395
      Width           =   495
   End
   Begin VB.TextBox txtCSSB00 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   116
      Text            =   "12,25"
      Top             =   5835
      Width           =   495
   End
   Begin VB.TextBox txtCSSB11 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   115
      Text            =   "12,25"
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox txtCSSA10 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   114
      Text            =   "12,25"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtCSSM01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   113
      Text            =   "12,25"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox txtCSSBJ 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   112
      Text            =   "12,25"
      Top             =   7275
      Width           =   495
   End
   Begin VB.TextBox txtCSSQ01 
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   111
      Text            =   "12,25"
      Top             =   8715
      Width           =   495
   End
   Begin VB.TextBox txtCSSE10 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   6720
      TabIndex        =   110
      Text            =   "12,25"
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox txtCSSC5 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   10320
      TabIndex        =   109
      Text            =   "12,25"
      Top             =   2955
      Width           =   495
   End
   Begin VB.TextBox txtCSSC1 
      BackColor       =   &H80000003&
      Height          =   285
      Left            =   600
      TabIndex        =   108
      Text            =   "12,25"
      Top             =   2950
      Width           =   495
   End
   Begin VB.TextBox txtCSC2 
      Height          =   285
      Left            =   3720
      TabIndex        =   97
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtCSC3 
      Height          =   285
      Left            =   6960
      TabIndex        =   96
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtCSC5 
      Height          =   285
      Left            =   10320
      TabIndex        =   95
      Text            =   "12,25"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtCSS01 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   94
      Text            =   "12,25"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtCSB11 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   93
      Text            =   "12,25"
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox txtCSA10 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   92
      Text            =   "12,25"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtCSQ01 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   91
      Text            =   "12,25"
      Top             =   9120
      Width           =   495
   End
   Begin VB.TextBox txtCSM01 
      Height          =   285
      Index           =   0
      Left            =   6600
      TabIndex        =   90
      Text            =   "12,25"
      Top             =   4920
      Width           =   495
   End
   Begin VB.TextBox txtCSBJ 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   89
      Text            =   "12,25"
      Top             =   7680
      Width           =   495
   End
   Begin VB.TextBox txtCSB00 
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   88
      Text            =   "12,25"
      Top             =   6240
      Width           =   495
   End
   Begin VB.TextBox txtCSE10 
      Height          =   285
      Left            =   6720
      TabIndex        =   87
      Text            =   "12,25"
      Top             =   6960
      Width           =   495
   End
   Begin VB.TextBox txtCSC1 
      Height          =   285
      Left            =   600
      TabIndex        =   86
      Text            =   "12,25"
      Top             =   3350
      Width           =   495
   End
   Begin VB.TextBox txtSB00 
      Height          =   285
      Left            =   480
      TabIndex        =   85
      Text            =   "e/kg"
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox txtSS01 
      Height          =   285
      Left            =   480
      TabIndex        =   84
      Text            =   "e/kg"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtSB11 
      Height          =   285
      Left            =   3600
      TabIndex        =   83
      Text            =   "e/kg"
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox txtSA10 
      Height          =   285
      Left            =   3600
      TabIndex        =   82
      Text            =   "e/kg"
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox txtSQ01 
      Height          =   285
      Left            =   480
      TabIndex        =   81
      Text            =   "e/kg"
      Top             =   8400
      Width           =   375
   End
   Begin VB.TextBox txtSM01 
      Height          =   285
      Left            =   6720
      TabIndex        =   80
      Text            =   "e/kg"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox txtSBJ 
      Height          =   285
      Left            =   480
      TabIndex        =   79
      Text            =   "e/kg"
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox txtSC5 
      Height          =   285
      Left            =   10200
      TabIndex        =   78
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtSC3 
      Height          =   285
      Left            =   6960
      TabIndex        =   77
      Text            =   "e/kg"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtSC2 
      Height          =   285
      Left            =   3720
      TabIndex        =   76
      Text            =   "e/kg"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtSC1 
      Height          =   285
      Left            =   480
      TabIndex        =   75
      Text            =   "e/kg"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox txtSE10 
      Height          =   285
      Left            =   6720
      TabIndex        =   74
      Text            =   "e/kg"
      Top             =   6000
      Width           =   375
   End
   Begin VB.TextBox txtEHQ01 
      Height          =   285
      Left            =   1440
      TabIndex        =   73
      Text            =   "e/h"
      Top             =   8115
      Width           =   735
   End
   Begin VB.TextBox txtQ01 
      Height          =   285
      Left            =   120
      TabIndex        =   72
      Text            =   "Squadratrice"
      Top             =   8115
      Width           =   855
   End
   Begin VB.TextBox txtEKQ01 
      Height          =   285
      Left            =   2160
      TabIndex        =   71
      Text            =   "e/kg"
      Top             =   8115
      Width           =   495
   End
   Begin VB.TextBox txtKHQ01 
      Height          =   285
      Left            =   960
      TabIndex        =   70
      Text            =   "kg/h"
      Top             =   8115
      Width           =   495
   End
   Begin VB.TextBox txtEHA10 
      Height          =   285
      Left            =   4440
      TabIndex        =   69
      Text            =   "e/h"
      Top             =   3795
      Width           =   735
   End
   Begin VB.TextBox txtA10 
      Height          =   285
      Left            =   3120
      TabIndex        =   68
      Text            =   "Accoppiatrice"
      Top             =   3795
      Width           =   855
   End
   Begin VB.TextBox txtEKA10 
      Height          =   285
      Left            =   5160
      TabIndex        =   67
      Text            =   "e/kg"
      Top             =   3795
      Width           =   495
   End
   Begin VB.TextBox txtKHA10 
      Height          =   285
      Left            =   3960
      TabIndex        =   66
      Text            =   "kg/h"
      Top             =   3795
      Width           =   495
   End
   Begin VB.TextBox txtEHM01 
      Height          =   285
      Left            =   7560
      TabIndex        =   65
      Text            =   "e/h"
      Top             =   3795
      Width           =   735
   End
   Begin VB.TextBox txtM01 
      Height          =   285
      Left            =   6240
      TabIndex        =   64
      Text            =   "Metallizzatore"
      Top             =   3795
      Width           =   855
   End
   Begin VB.TextBox txtEKM01 
      Height          =   285
      Left            =   8280
      TabIndex        =   63
      Text            =   "e/kg"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtKHM01 
      Height          =   285
      Left            =   7080
      TabIndex        =   62
      Text            =   "kg/h"
      Top             =   3795
      Width           =   495
   End
   Begin VB.TextBox txtEHE10 
      Height          =   285
      Left            =   7680
      TabIndex        =   61
      Text            =   "e/h"
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtE10 
      Height          =   285
      Left            =   6120
      TabIndex        =   60
      Text            =   "Estrusore"
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtEKE10 
      Height          =   285
      Left            =   8400
      TabIndex        =   59
      Text            =   "e/kg"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox txtKHE10 
      Height          =   285
      Left            =   7080
      TabIndex        =   58
      Text            =   "kg/h"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox txtEHB11 
      Height          =   285
      Left            =   4680
      TabIndex        =   57
      Text            =   "e/h"
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox txtB11 
      Height          =   285
      Left            =   3360
      TabIndex        =   56
      Text            =   "t. Sleeve"
      Top             =   7200
      Width           =   855
   End
   Begin VB.TextBox txtEKB11 
      Height          =   285
      Left            =   5400
      TabIndex        =   55
      Text            =   "e/kg"
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox txtKHB11 
      Height          =   285
      Left            =   4200
      TabIndex        =   54
      Text            =   "kg/h"
      Top             =   7200
      Width           =   495
   End
   Begin VB.TextBox txtEHS01 
      Height          =   285
      Left            =   1440
      TabIndex        =   53
      Text            =   "e/h"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtS01 
      Height          =   285
      Left            =   120
      TabIndex        =   52
      Text            =   "Sleeve"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtEKS01 
      Height          =   285
      Left            =   2160
      TabIndex        =   51
      Text            =   "e/kg"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtKHS01 
      Height          =   285
      Left            =   960
      TabIndex        =   50
      Text            =   "kg/h"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txtEHB00 
      Height          =   285
      Left            =   1440
      TabIndex        =   49
      Text            =   "e/h"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtB00 
      Height          =   285
      Left            =   120
      TabIndex        =   48
      Text            =   "B00"
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtEKB00 
      Height          =   285
      Left            =   2160
      TabIndex        =   47
      Text            =   "e/kg"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtKHB00 
      Height          =   285
      Left            =   960
      TabIndex        =   46
      Text            =   "kg/h"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtEHC1 
      Height          =   285
      Left            =   1440
      TabIndex        =   45
      Text            =   "247,56"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtEHC2 
      Height          =   285
      Left            =   4680
      TabIndex        =   44
      Text            =   "e/h"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtEHC3 
      Height          =   285
      Left            =   7920
      TabIndex        =   43
      Text            =   "e/h"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtEHC5 
      Height          =   285
      Left            =   11160
      TabIndex        =   42
      Text            =   "e/h"
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox txtEHBJ 
      Height          =   285
      Left            =   1440
      TabIndex        =   41
      Text            =   "e/h"
      Top             =   6645
      Width           =   735
   End
   Begin VB.TextBox txtKHC1 
      Height          =   285
      Left            =   960
      TabIndex        =   40
      Text            =   "758"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtEKC1 
      Height          =   285
      Left            =   2160
      TabIndex        =   39
      Text            =   "12,25"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtKHC2 
      Height          =   285
      Left            =   4200
      TabIndex        =   38
      Text            =   "kg/h"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtEKC2 
      Height          =   285
      Left            =   5400
      TabIndex        =   37
      Text            =   "e/kg"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtC1 
      Height          =   285
      Left            =   120
      TabIndex        =   36
      Text            =   "Calandra1"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtC2 
      Height          =   285
      Left            =   3360
      TabIndex        =   35
      Text            =   "Calandra2"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtC3 
      Height          =   285
      Left            =   6600
      TabIndex        =   34
      Text            =   "Calandra3"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtEKC3 
      Height          =   285
      Left            =   8640
      TabIndex        =   33
      Text            =   "e/kg"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtKHC3 
      Height          =   285
      Left            =   7440
      TabIndex        =   32
      Text            =   "kg/h"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtC5 
      Height          =   285
      Left            =   9840
      TabIndex        =   31
      Text            =   "Calandra5"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtEKC5 
      Height          =   285
      Left            =   11880
      TabIndex        =   30
      Text            =   "e/kg"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtKHC5 
      Height          =   285
      Left            =   10680
      TabIndex        =   29
      Text            =   "kg/h"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtBJ 
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Text            =   "jumbo"
      Top             =   6645
      Width           =   855
   End
   Begin VB.TextBox txtEKBJ 
      Height          =   285
      Left            =   2160
      TabIndex        =   27
      Text            =   "e/kg"
      Top             =   6645
      Width           =   495
   End
   Begin VB.TextBox txtKHBJ 
      Height          =   285
      Left            =   960
      TabIndex        =   26
      Text            =   "kg/h"
      Top             =   6645
      Width           =   495
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   9
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "-0.18"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   8
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "1.33"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   7
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0.91"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   6
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0.66"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   ""
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   5
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "1500"
      Top             =   960
      Width           =   495
   End
   Begin MSComctlLib.ImageList ilToolBar 
      Left            =   12000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":0E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":136A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":1DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":2336
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATC.frx":287A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   3
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   ""
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   4
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "200"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "H039500"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtMescola 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "ECB29"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtVenditore 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "00530.1"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdDettagli 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Dettagli"
      Height          =   615
      Left            =   11280
      MousePointer    =   1  'Arrow
      Picture         =   "frmATC.frx":2DBE
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.Toolbar tbCommand 
      Align           =   2  'Align Bottom
      Height          =   630
      Left            =   0
      TabIndex        =   12
      Top             =   9675
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   1111
      ButtonWidth     =   1349
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Nuovo"
            Key             =   "N"
            Object.ToolTipText     =   "Nuovo Reclamo"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "prova"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Salva"
            Key             =   "S"
            Object.ToolTipText     =   "Salva Reclamo Corrente"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Annulla"
            Key             =   "A"
            Object.ToolTipText     =   "Annulla modifiche su reclamo corrente"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modifica"
            Key             =   "M"
            Object.ToolTipText     =   "Modifica il Reclamo Corrente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ricalcola"
            Key             =   "C"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   300
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stampa"
            Key             =   "P"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "phRight"
            Style           =   4
            Object.Width           =   500
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esci"
            Key             =   "X"
            Object.ToolTipText     =   "Esci"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MA514"
      Height          =   195
      Left            =   10560
      TabIndex        =   352
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crr"
      Height          =   195
      Left            =   10200
      TabIndex        =   351
      Top             =   1320
      Width           =   195
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Std"
      Height          =   195
      Left            =   10200
      TabIndex        =   350
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label36 
      Caption         =   "Totale"
      Height          =   255
      Left            =   2160
      TabIndex        =   345
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label35 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   840
      TabIndex        =   344
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label26 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   120
      TabIndex        =   343
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   "Std"
      Height          =   255
      Left            =   120
      TabIndex        =   342
      Top             =   8640
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "Crr"
      Height          =   255
      Left            =   120
      TabIndex        =   341
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Macchina"
      Height          =   255
      Left            =   240
      TabIndex        =   340
      Top             =   2040
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   0
      X2              =   12500
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label65 
      Caption         =   "Crr"
      Height          =   255
      Left            =   120
      TabIndex        =   213
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label64 
      Caption         =   "Std"
      Height          =   255
      Left            =   120
      TabIndex        =   212
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label63 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   120
      TabIndex        =   211
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label60 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   840
      TabIndex        =   205
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label59 
      Caption         =   "Totale"
      Height          =   255
      Left            =   2160
      TabIndex        =   204
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label56 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   840
      TabIndex        =   193
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label55 
      Caption         =   "Totale"
      Height          =   255
      Left            =   2160
      TabIndex        =   192
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label Label54 
      Caption         =   "Crr"
      Height          =   255
      Left            =   120
      TabIndex        =   186
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label53 
      Caption         =   "Std"
      Height          =   255
      Left            =   120
      TabIndex        =   185
      Top             =   5760
      Width           =   375
   End
   Begin VB.Label Label52 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   120
      TabIndex        =   184
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label51 
      Caption         =   "Kg/h"
      Height          =   255
      Left            =   10680
      TabIndex        =   183
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label50 
      Caption         =   "€/Kg"
      Height          =   255
      Left            =   12000
      TabIndex        =   182
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label49 
      Caption         =   "€/h"
      Height          =   255
      Left            =   11400
      TabIndex        =   181
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label48 
      Caption         =   "Kg/h"
      Height          =   255
      Left            =   7440
      TabIndex        =   180
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label47 
      Caption         =   "€/Kg"
      Height          =   255
      Left            =   8760
      TabIndex        =   179
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label46 
      Caption         =   "€/h"
      Height          =   255
      Left            =   8160
      TabIndex        =   178
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label45 
      Caption         =   "Kg/h"
      Height          =   255
      Left            =   4200
      TabIndex        =   177
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label44 
      Caption         =   "€/Kg"
      Height          =   255
      Left            =   5520
      TabIndex        =   176
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label43 
      Caption         =   "€/h"
      Height          =   255
      Left            =   4920
      TabIndex        =   175
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblLT 
      Caption         =   "Totale"
      Height          =   255
      Left            =   2160
      TabIndex        =   164
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblLR 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   840
      TabIndex        =   163
      Top             =   4080
      Width           =   495
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   12500
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Label Label34 
      Caption         =   "Totale"
      Height          =   255
      Left            =   11880
      TabIndex        =   152
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label33 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   10680
      TabIndex        =   151
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label32 
      Caption         =   "Totale"
      Height          =   255
      Left            =   8640
      TabIndex        =   145
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label31 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   7440
      TabIndex        =   144
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label30 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   3360
      TabIndex        =   138
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label29 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   6600
      TabIndex        =   137
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label28 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   9840
      TabIndex        =   136
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblLS 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   120
      TabIndex        =   135
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label Label25 
      Caption         =   "Totale"
      Height          =   255
      Left            =   5400
      TabIndex        =   134
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label24 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   4200
      TabIndex        =   133
      Top             =   2640
      Width           =   495
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   12500
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   12500
      Y1              =   3690
      Y2              =   3690
   End
   Begin VB.Label yg 
      Caption         =   "%Rec"
      Height          =   255
      Left            =   840
      TabIndex        =   126
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblLSt 
      Caption         =   "Std"
      Height          =   255
      Left            =   120
      TabIndex        =   125
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label lblLCr 
      Caption         =   "Crr"
      Height          =   255
      Left            =   120
      TabIndex        =   124
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblC2 
      Caption         =   "Macchina"
      Height          =   255
      Left            =   3480
      TabIndex        =   107
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblC3 
      Caption         =   "Macchina"
      Height          =   255
      Left            =   6720
      TabIndex        =   106
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblC5 
      Caption         =   "Macchina"
      Height          =   255
      Left            =   9960
      TabIndex        =   105
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label kj 
      Caption         =   "Totale"
      Height          =   255
      Left            =   2160
      TabIndex        =   104
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Crr"
      Height          =   255
      Left            =   120
      TabIndex        =   103
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label ih 
      Caption         =   "Std"
      Height          =   255
      Left            =   120
      TabIndex        =   102
      Top             =   2955
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "%Sc"
      Height          =   255
      Left            =   120
      TabIndex        =   101
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "€/h"
      Height          =   255
      Left            =   1680
      TabIndex        =   100
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "€/Kg"
      Height          =   255
      Left            =   2280
      TabIndex        =   99
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lblC1 
      Caption         =   "Kg/h"
      Height          =   255
      Left            =   960
      TabIndex        =   98
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kg MA514"
      Height          =   195
      Left            =   5760
      TabIndex        =   25
      Top             =   720
      Width           =   750
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Crr"
      Height          =   195
      Left            =   8520
      TabIndex        =   24
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrizione prodotto"
      Height          =   195
      Left            =   1560
      TabIndex        =   23
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   4440
      TabIndex        =   22
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prezzi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6600
      TabIndex        =   21
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8160
      TabIndex        =   20
      Top             =   840
      Width           =   120
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Std"
      Height          =   195
      Left            =   7680
      TabIndex        =   19
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrizione mescola"
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   12375
   End
   Begin VB.Label lblTitolo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Analisi dei Costi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   630
      Left            =   4680
      TabIndex        =   6
      Top             =   0
      Width           =   3570
   End
   Begin VB.Image iLogo 
      Height          =   450
      Left            =   120
      Picture         =   "frmATC.frx":32F0
      Top             =   0
      Width           =   450
   End
   Begin VB.Label lblReclamoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Range spessore"
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prodotto finito"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mescola"
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9000
      TabIndex        =   0
      Top             =   960
      Width           =   165
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   7320
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2775
   End
   Begin VB.Shape shpC2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2020
      Width           =   2775
   End
   Begin VB.Shape shpC3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   6480
      Shape           =   4  'Rounded Rectangle
      Top             =   2020
      Width           =   2775
   End
   Begin VB.Shape shpC5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   9720
      Shape           =   4  'Rounded Rectangle
      Top             =   2020
      Width           =   2775
   End
   Begin VB.Shape shpC1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   7575
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   2025
      Width           =   2775
   End
End
Attribute VB_Name = "frmCosti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mKey As String, rsATC As Recordset
Private fNuovoRecord As Boolean, sStato As String
Private mStatus As Integer
Private Enum CMDS
   cNuovo = 1
   cSalva = 3
   cAnnulla = 5
   cModifica = 7
   cChiudi = 9
   cStampa = 11
   cEsci = 13
End Enum
Private Enum Status
   cNew = 0
   cMod = 1
   cExists = 2
   cNotExists = 3
End Enum
Dim sqlc As String, i As Integer
Dim rsCiclo As Recordset
Dim dPrezzoCrr As Double, dPrezzoStd As Double
Dim f1 As Boolean, f2 As Boolean, f3 As Boolean, f5 As Boolean, fE As Boolean
Dim fS As Boolean, fM As Boolean, fA As Boolean, fAcc As Boolean
Dim fRicalcola As Boolean, c As Control
Dim sScarto As String, sRecupero As String
Dim iLinea As Integer

Public Property Let Status(vValue As Integer)
   If vValue >= cNew And vValue <= cNotExists Then
      mStatus = vValue
   End If
   With tbCommand
    If sFlag = "ATC" Then ' controllo che non possano modificare i dati dal MKG
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = True
            .Buttons(cAnnulla).Enabled = True
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = True
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = False
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = True
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = True
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
      End Select
    Else
      Select Case mStatus
         Case cNew, cMod
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
         Case cNotExists
            .Buttons(cNuovo).Enabled = False
            .Buttons(cSalva).Enabled = False
            .Buttons(cAnnulla).Enabled = False
            .Buttons(cModifica).Enabled = False
            .Buttons(cChiudi).Enabled = False
            .Buttons(cStampa).Enabled = True
            .Buttons(cEsci).Enabled = True
      End Select
    End If
    
   End With
'Private Sub NuovoReclamo()
'End Sub
End Property
Public Property Get Status() As Integer
   Status = mStatus
End Property

Public Property Let Key(vValue As String)
Dim c As Control
    mKey = Trim$(vValue)
'    CercaReclamoVenditore
    CercaMescola
'    CercaReclamoATC
    CercaProduzione
    For Each c In Me.Controls
        If TypeOf c Is TextBox Then
 '           If c.Visible = True Then
                c.Locked = True
 '           End If
'               Debug.Print c.Name
'               Debug.Print c.Visible
        End If
    Next

'    Set datElenco.Recordset = rsATC
'    datElenco.RecordSource = sqlc
'    txtODLInputTest.Text = mKey
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Private Sub Rete()
   Dim lDiff As Long
    Load frmRete
    With frmRete
        .Top = Me.Top + (Me.Height - frmRete.Height)
        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
        lDiff = Screen.Width * 0.98 - (.Left + .Width)
        If lDiff < 0 Then
            .Left = .Left + lDiff
        End If
'        .Key = txtNumeroRisposta.Text
        .Show vbModal
    End With
End Sub

Private Sub cmdDettagli_Click()
    Dim lDiff As Long
    Load frmMFG
    With frmMFG
        .Top = Me.Top
        .Left = Me.Left + Me.Width
'   Controllo di non superare le dimensioni dello schermo
'       altrimenti la sposto indietro
        lDiff = Screen.Width * 0.98 - (.Left + .Width)
        If lDiff < 0 Then
            .Left = .Left + lDiff
        End If
        .Key = Me.Key
        .Show
    End With
End Sub




Private Sub SalvaRec()
'   Dim sCod As Variant
'   If fNuovoRecord Then
'      rsATC.AddNew
'      txtNumeroRisposta.Text = CalcUltimaRisposta()
'      rsATC("NumeroRisposta") = txtNumeroRisposta.Text
'      rsATC("ODL") = Me.Key
'   Else
'      rsATC.Edit
'   End If
''   If Not IsNull(txtNumeroRisposta.Text) Then
''        rsATC("NumeroRisposta") = txtNumeroRisposta.Text
''   Else
''         rsATC("NumeroRisposta") = Null
''   End If
'   If Not (IsNull(txtDataRisposta.Text) Or txtDataRisposta.Text = "") Then
'      rsATC("DataRisposta") = txtDataRisposta.Text
'   Else
'       rsATC("DataRisposta") = Null
'   End If
'   sCod = SalvaCodiceDifetto(cmbDifetto.Text)
'   lblCodDifetto.Caption = sCod
'   If Not (IsNull(lblCodDifetto.Caption) Or lblCodDifetto.Caption = "") Then
'       rsATC("CodDifetto") = lblCodDifetto.Caption
'   Else
'       rsATC("CodDifetto") = Null
'   End If
'   If Not (IsNull(txtQtaAccettata.Text) Or txtQtaAccettata.Text = "") Then
'        rsATC("QtaAccettata") = Val(txtQtaAccettata.Text)
'   Else
'         rsATC("QtaAccettata") = Null
'   End If
'
''     Carico i dati del codfamiglia dall'rs
'   sCod = SalvaCodiceFamiglia(cmbFamiglia.Text)
 '  lblCodFamiglia = sCod
'   If Not (IsNull(lblCodFamiglia.Caption) Or lblCodFamiglia.Caption = "") Then
'       rsATC("CodFamiglia") = lblCodFamiglia
'   Else
'       rsATC("CodFamiglia") = Null
'   End If
'   If Not (IsNull(txtRisposta.Text) Or txtRisposta.Text = "") Then
'       rsATC("Risposta") = txtRisposta
'   Else
'       rsATC("Risposta") = Null
'   End If
'   If Not (IsNull(txtAzioni.Text) Or txtAzioni.Text = "") Then
'       rsATC("AzioniCorrettive") = txtAzioni.Text
'   Else
 '      rsATC("AzioniCorrettive") = Null
 '  End If
'   If optIq(0).Value Then
'       rsATC("IndiceQualita") = True
'   Else
'       If optIq(1).Value Then
'           rsATC("IndiceQualita") = False
'       Else
'           rsATC("IndiceQualita") = Null
'       End If
'   End If
'   If optExt(0).Value Then
'       rsATC("Esterno") = True
'   Else
'       If optExt(1).Value Then
'           rsATC("Esterno") = False
'       Else
'           rsATC("Esterno") = Null
'       End If
'   End If
'   If optSol(0).Value Then
'       rsATC("Soluzione") = 0
'   Else
'       If optSol(1).Value Then
'           rsATC("Soluzione") = 1
'       Else
'           If optSol(2).Value Then
'              rsATC("Soluzione") = 2
'           Else
'   '            If optSol(3).Value Then
'    '               rsATC("Soluzione") = 3
'     '          Else
'                   rsATC("Soluzione") = Null
 '      '        End If
 '         End If
 '      End If
 '  End If
   
'   rsATC.Update
'   CercaReclamoATC
''     disabilito i campi
'   Campi False
'   fNuovoRecord = False
End Sub

Private Sub Form_Load()
   Dim btn As Button, c As Control
   
'  Carica i combo
'   ListCodiceDifetto
'   ListCodiceFamiglia
   fRicalcola = False
'   Campi False
   Macchine False
    'Nascondo tutti i pulsanti
'    cmdNew.Visible = False
'    cmdSalva.Visible = False
'    cmdAnnulla.Visible = False
'    cmdModifica.Visible = False
   For Each btn In tbCommand.Buttons
     If btn.Key <> "X" Then
        btn.Enabled = False
        
     End If
   Next
    fAcc = False
    
        
End Sub

Private Sub Form_Resize()
   Dim b As Button, l As Long, idx As Integer
   If Me.WindowState <> vbMinimized Then
'      l = shpRisposta.Left * 2 + shpRisposta.Width + Me.Width - Me.ScaleWidth
      If Me.Width < l Then
         Me.Width = l
      End If
      l = 0
      Me.Refresh
      For Each b In tbCommand.Buttons
         If b.Caption <> "phRight" Then
            l = l + b.Width
         Else
            idx = b.Index
         End If
      Next
      Set b = tbCommand.Buttons(idx)
      b.Width = Me.ScaleWidth - l
      Set b = Nothing
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Load frmVenditori
'    frmVenditori.Show
'   rsATC.Close
'   Set rsATC = Nothing
End Sub

Private Sub CercaReclamoVenditore()
   Dim rsVenditore As Recordset, sqlc As String, i As Integer
   
   
   sqlc = "SELECT ODL, DataReclamo, QtaReclamata, " & _
           "CausaReclamo, Venditore, NumeroReclamo, Stato " & _
           "FROM ReclamoVenditore " & _
           "where ODL='" & Me.Key & "'"
   
   
   
   
   Set rsVenditore = db.OpenRecordset(sqlc, dbOpenDynaset)
   If Not rsVenditore.EOF Then
     If sFlag = "ATC" Then ' Controllo chi apre il DB
'      If rsVenditore("Stato") = "N" Then
'         rsVenditore.Edit
'         rsVenditore("Stato") = "V"
'         rsVenditore.Update
'         frmTV.ChangedTV = True
      'End If
     End If
      For i = 0 To rsVenditore.Fields.Count - 1
         If Not IsNull(rsVenditore(i)) Then
             txtVenditore(i).Text = rsVenditore(i)
         Else
             txtVenditore(i).Text = ""
         End If
      Next i
      sStato = rsVenditore("Stato")
   Else
      Err.Raise 13, , "Errore nel data source"
   End If
   rsVenditore.Close
   Set rsVenditore = Nothing

End Sub
Private Sub CercaReclamoATC()
    Dim sqlc As String, i As Integer
   
    sqlc = "SELECT ODL, NumeroRisposta, DataRisposta, " & _
            "Risposta, AzioniCorrettive, CodFamiglia, " & _
            "CodDifetto, Esterno, StatoReclamo, IndiceQualita, " & _
            "Soluzione, QtaAccettata FROM ATC where ODL='" & Me.Key & "'"
    Set rsATC = db.OpenRecordset(sqlc, dbOpenDynaset)
    
   If Not rsATC.EOF Then
      FoundODL
      Me.Status = cExists
    Else
      Me.Status = cNotExists
    End If
    'For i = 0 To rsATC.Fields.Count - 1
    '    txtReclami(i).DataField = rsATC.Fields(i).Name
    'Next i
End Sub
Private Sub FoundODL()
'   Dim iSoluz As Integer
'    If Not IsNull(rsATC("NumeroRisposta")) Then
'        txtNumeroRisposta = rsATC("NumeroRisposta")
'    Else
'        txtNumeroRisposta = ""
'    End If
'    If Not IsNull(rsATC("DataRisposta")) Then
'        txtDataRisposta = rsATC("DataRisposta")
'    Else
'        txtDataRisposta = ""
'    End If
'    If Not IsNull(rsATC("CodDifetto")) Then
'        lblCodDifetto = rsATC("CodDifetto")
'        cmbDifetto.Text = FoundCodiceDifetto(lblCodDifetto.Caption)
'    Else
'        lblCodDifetto = ""
'        cmbDifetto.ListIndex = -1
'    End If
'    If Not IsNull(rsATC("CodFamiglia")) Then
'      lblCodFamiglia = rsATC("CodFamiglia")
'      cmbFamiglia.Text = FoundCodiceFamiglia(lblCodFamiglia.Caption)
'    Else
'     lblCodFamiglia = ""
'     cmbFamiglia.ListIndex = -1
'    End If
'    If Not IsNull(rsATC("Risposta")) Then
 '     txtRisposta = rsATC("Risposta")
'    Else
'        txtRisposta = ""
'    End If
'    If Not IsNull(rsATC("AzioniCorrettive")) Then
'        txtAzioni = rsATC("AzioniCorrettive")
'    Else
'        txtAzioni = ""
'    End If
 '   If Not IsNull(rsATC("IndiceQualita")) Then
'        If rsATC("IndiceQualita") Then
'            optIq(0).Value = True
'        Else
'            optIq(1).Value = True
'        End If
'    Else
'          optIq(0).Value = False
'          optIq(1).Value = False
'    End If
'    If Not IsNull(rsATC("Esterno")) Then
'        If rsATC("Esterno") Then
'            optExt(0).Value = True
'        Else
'            optExt(1).Value = True
'        End If
'    Else
'        optExt(0).Value = False
'        optExt(1).Value = False
'    End If
'    If Not IsNull(rsATC("Soluzione")) Then
'      iSoluz = rsATC("Soluzione")
'      optSol(iSoluz).Value = True
'    Else
'        optSol(0).Value = False
'        optSol(1).Value = False
'        optSol(2).Value = False
'    End If
'   If Not IsNull(rsATC("QtaAccettata")) Then
'      txtQtaAccettata = rsATC("QtaAccettata")
'    Else
'        txtQtaAccettata = ""
'    End If
End Sub
Private Sub ListCodiceDifetto()
   Dim rsDifetto As Recordset, sqlc As String, sMsg As String
   sqlc = "SELECT DescrizioneDifetto " & _
           "FROM Difetto "
 '  Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
 '  cmbDifetto.Clear
 '  Do While Not rsDifetto.EOF
 '     cmbDifetto.AddItem rsDifetto(0)
 '     rsDifetto.MoveNext
 '  Loop
 '  rsDifetto.Close
 '  Set rsDifetto = Nothing
 '  cmbDifetto.ListIndex = -1
End Sub
Private Sub ListCodiceFamiglia()
'    Dim rsFamiglia As Recordset, sqlc As String, sMsg As String
'    sqlc = "SELECT DescrizioneFamiglia " & _
'            "FROM Famiglia order by 1"
'    Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
'    cmbFamiglia.Clear
'    Do While Not rsFamiglia.EOF
'        cmbFamiglia.AddItem rsFamiglia(0)
'        rsFamiglia.MoveNext
'    Loop
'    rsFamiglia.Close
'    Set rsFamiglia = Nothing
'    cmbFamiglia.ListIndex = -1
    
End Sub
Private Function FoundCodiceDifetto(cod As String)
'   Dim rsDifetto As Recordset, sqlc As String
'   sqlc = "SELECT DescrizioneDifetto " & _
'           "FROM Difetto where CodDifetto='" & cod & "'"
'   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   If Not rsDifetto.EOF Then
'       FoundCodiceDifetto = rsDifetto("DescrizioneDifetto")
'   Else
'       FoundCodiceDifetto = ""
''   End If
'   rsDifetto.Close
'   Set rsDifetto = Nothing
End Function
Private Function FoundCodiceFamiglia(cod As String)
'   Dim rsFamiglia As Recordset, sqlc As String
'   sqlc = "SELECT DescrizioneFamiglia " & _
'           "FROM Famiglia where CodFamiglia='" & cod & "'"
'   Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   If Not rsFamiglia.EOF Then
'       FoundCodiceFamiglia = rsFamiglia("DescrizioneFamiglia")
'   Else
'       FoundCodiceFamiglia = ""
'   End If
'   rsFamiglia.Close
'   Set rsFamiglia = Nothing
End Function
Private Function SalvaCodiceDifetto(s As String)
'   Dim rsDifetto As Recordset, sqlc As String, sDescrizione As String
'   Dim iPos As Integer
'   iPos = InStr(s, "'")
'   If iPos <> 0 Then
'      s = Left$(s, iPos) & "'" & Right$(s, Len(s) - iPos)
'
'   End If
'   sqlc = "SELECT CodDifetto " & _
'           "FROM Difetto where DescrizioneDifetto='" & Trim$(s) & "'"
'   Set rsDifetto = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   If Not rsDifetto.EOF Then
'      SalvaCodiceDifetto = rsDifetto("CodDifetto")
'   Else
'      SalvaCodiceDifetto = ""
'   End If
'   rsDifetto.Close
'   Set rsDifetto = Nothing
End Function

Private Function SalvaCodiceFamiglia(s As String)
'   Dim rsFamiglia As Recordset, sqlc As String, sDescrizione As String
'
'   sqlc = "SELECT CodFamiglia " & _
'           "FROM Famiglia where DescrizioneFamiglia='" & s & "'"
'   Set rsFamiglia = db.OpenRecordset(sqlc, dbOpenSnapshot)
'   If Not rsFamiglia.EOF Then
'       SalvaCodiceFamiglia = rsFamiglia("CodFamiglia")
'   Else
'      SalvaCodiceFamiglia = ""
'   End If
'   rsFamiglia.Close
'   Set rsFamiglia = Nothing
End Function

Private Sub Campi(Flag As Boolean)
'    txtNumeroRisposta.Enabled = False
'    txtDataRisposta.Enabled = Flag
'    txtRisposta.Enabled = Flag
'    txtAzioni.Enabled = Flag
'    fraIQ.Enabled = Flag
'    fraEsterno.Enabled = Flag
'    fraSoluzione.Enabled = Flag
'    cmbDifetto.Enabled = Flag
'    cmbFamiglia.Enabled = Flag
'    txtQtaAccettata.Enabled = Flag
'    txtNumeroRisposta.Locked = Not Flag
'    txtDataRisposta.Locked = Not Flag
'    txtRisposta.Locked = Not Flag
'    txtAzioni.Locked = Not Flag
'    txtQtaAccettata.Locked = Not Flag
End Sub

Private Sub CancellaCampi()
'   txtNumeroRisposta.Text = ""
'   txtDataRisposta.Text = ""
'   txtRisposta.Text = ""
'   txtAzioni.Text = ""
'   txtQtaAccettata.Text = ""
'   optIq(0).Value = False
'   optIq(1).Value = False
'
'   optExt(0).Value = False
'   optExt(1).Value = False
'
'   optSol(0).Value = False
'   optSol(1).Value = False
'   optSol(2).Value = False
'
'   cmbDifetto.ListIndex = -1
'   cmbFamiglia.ListIndex = -1
End Sub


Private Sub tbCommand_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim iRes As Integer, rsTmp As Recordset, sqlc As String
   Select Case Button.Key
      Case "X"
         Unload Me
      Case "N"
         fNuovoRecord = True
         Campi True
         CancellaCampi
         Me.Status = cNew
      Case "S"
         SalvaRec
         Me.Status = cExists
      Case "A"
        fRicalcola = False
        txtSC1.Text = "e/kg"
        txtRC1.Text = "e/kg"
        txtSC2.Text = "e/kg"
        txtRC2.Text = "e/kg"
        txtSC3.Text = "e/kg"
        txtRC3.Text = "e/kg"
        txtSC5.Text = "e/kg"
        txtRC5.Text = "e/kg"
        txtSB00.Text = "e/kg"
        txtRB00.Text = "e/kg"
        txtSBJ.Text = "e/kg"
        txtRBJ.Text = "e/kg"
        txtSB11.Text = "e/kg"
        txtRB11.Text = "e/kg"
        txtSS01.Text = "e/kg"
        txtRS01.Text = "e/kg"
        txtSA10.Text = "e/kg"
        txtRA10.Text = "e/kg"
        txtSE10.Text = "e/kg"
        txtRE10.Text = "e/kg"
        txtSQ01.Text = "e/kg"
        txtRQ01.Text = "e/kg"
        txtSM01.Text = "e/kg"
        txtRM01.Text = "e/kg"

        CercaProduzione
         
        For Each c In Me.Controls
            If TypeOf c Is TextBox Then
                c.Locked = True
            End If
        Next
        Me.Status = cExists
      Case "M"
        For Each c In Me.Controls
            If TypeOf c Is TextBox Then
                c.Locked = False
            End If
        Next
        txtMescola(7).Enabled = True
        txtMescola(8).Enabled = True
        Me.Status = cMod
      Case "C"
        fRicalcola = True
        'Controllo che si inserisca un spessore per il PVC in caso di accoppiati
        If txtSP.Visible = True Then
            If IsNull(txtSP.Text) Or txtSP.Text = "" Then
                MsgBox "si deve inserire uno spessore per il PVC"
                txtSP.SetFocus
            Else
                fAcc = True
            End If
        End If
        CercaProduzione
        For Each c In Me.Controls
            If TypeOf c Is TextBox Then
                c.Locked = True
            End If
         Next
        fRicalcola = False


'         iRes = MsgBox("Vuoi chiudere il corrente reclamo", _
'               vbYesNo + vbDefaultButton2 + vbQuestion)
'         If iRes = vbYes Then
'            sqlc = "Update ReclamoVenditore set Stato = 'C' " & _
'                     "Where ODL = '" & txtVenditore(0) & "' and " & _
'                     "Stato = 'V'"
'            db.Execute (sqlc)
'            sqlc = "Update ATC set StatoReclamo = 'N' " & _
'                  "where ODL = '" & txtVenditore(0) & "'"
'            db.Execute (sqlc)
'            sStato = "C"
'            Me.Status = cExists
'            frmTV.ChangedTV = True
'         End If
      Case "P"
'         fNuovoRecord = False
'         Rete
         Me.PrintForm
         Me.Status = cExists
   End Select
End Sub
Private Function CalcUltimaRisposta()
   Dim rsT As Recordset, sqlc As String
   sqlc = "update UltimaRisposta " & _
               "set UltimaRisposta = UltimaRisposta + 1 " & _
               "Where dummy=0"
   db.Execute (sqlc)
   sqlc = "Select UltimaRisposta from UltimaRisposta " & _
               "Where dummy = 0"
   Set rsT = db.OpenRecordset(sqlc, dbOpenSnapshot)
   If Not rsT.EOF Then
      CalcUltimaRisposta = rsT(0)
   Else
      Err.Raise 555, , "Errore nel calcolo del numero risposta"
   End If
   rsT.Close
   Set rsT = Nothing
End Function

Private Sub CercaMescola()
   Dim rsMescola As Recordset, sqlc As String, i As Integer
    Dim sWhere As String, rsMa As Recordset, rsA As Recordset

Select Case Len(Me.Key)
    Case Is < 11
        sWhere = "WHERE Generale.Mescola='" & Trim$(Left$(Me.Key, 5)) & "' " & _
                    " and [LdP]='" & Trim$(Mid$(Me.Key, 6, 4)) & "' "
    Case Else
        sWhere = "WHERE Generale.Mescola='" & Trim$(Left$(Me.Key, 5)) & "' " & _
                    " and ([LdP] & 'A' & [ProdFin] & [desc])='" & _
                    Trim$(Mid$(Me.Key, 6, Len(Me.Key))) & "' "
End Select
    
    sqlc = "SELECT Generale.Mescola, Generale.ProdFin, Generale.Desc, " & _
             " Generale.DescMescola, Spessore.DA, " & _
             "Spessore.A, Generale.KgMa514, Generale.Standard, " & _
             "MescoleMaxCostoCRR.Corrente, Format([Corrente]-[Standard],'0.00') " & _
             "AS DMa514, MescoleMaxCostoCRR.MaxDiData AS Data " & _
         "FROM (Generale LEFT JOIN MescoleMaxCostoCRR ON " & _
             "Generale.Mescola = MescoleMaxCostoCRR.Mescola) " & _
             "LEFT JOIN Spessore ON Generale.Mescola = Spessore.Mesc " & _
          sWhere


   Set rsMescola = db.OpenRecordset(sqlc, dbOpenDynaset)
   If Not rsMescola.EOF Then
     If sFlag = "ATC" Then ' Controllo chi apre il DB
'      If rsVenditore("Stato") = "N" Then
'         rsVenditore.Edit
'         rsVenditore("Stato") = "V"
'         rsVenditore.Update
'         frmTV.ChangedTV = True
      'End If
     End If
      For i = 0 To rsMescola.Fields.Count - 1
         If Not IsNull(rsMescola(i)) Then
             txtMescola(i).Text = rsMescola(i)
         Else
             txtMescola(i).Text = ""
         End If
      Next i
  '    sStato = rsMescola("Stato")
   Else
      Err.Raise 13, , "Errore nel data source"
   End If
   rsMescola.Close
   Set rsMescola = Nothing

    sqlc = "SELECT Max(CMa.Data) AS MaxDiData, CMa.CStMa514, CMa.CCMa514 " & _
                "FROM CostoMa514 AS CMa " & _
                "GROUP BY CMa.CStMa514, CMa.CCMa514"

    Set rsMa = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsMa.EOF Then
        txtCStMa514.Text = rsMa("CStMa514")
        
    End If
    rsMa.Close
    Set rsMa = Nothing
'   Controllo la presenza degli accoppiati
    sqlc = "SELECT ProdFin, MP, Costo, Desc, Spessore " & _
            "From MP " & _
            "Where ProdFin = '" & Mid$(Me.Key, 11, 7) & " '"
    Set rsA = db.OpenRecordset(sqlc, dbOpenSnapshot)
    If Not rsA.EOF Then
            
            txtMP.Text = rsA("MP")
            txtMP.Visible = True
            txtMPD.Text = rsA("Desc")
            txtMPD.Visible = True
            txtMPC.Text = rsA("Costo")
            txtMPC.Visible = True
            txtMPS.Text = rsA("Spessore")
            txtMPS.Visible = True
            txtMPPS.Text = "0.93"
            txtMPPS.Visible = True
            txtSP.Text = ""
            txtSP.Visible = True
    End If
    
    rsA.Close
    Set rsA = Nothing
    
    
End Sub
Private Sub CercaProduzione()

f1 = False
f2 = False
f3 = False
f5 = False
fE = False
fS = False
fM = False
fA = False
    
    sqlc = "SELECT PM.Ciclo, PM.Ldp, PM.CdL, Scarti.Scarto, " & _
                "Scarti.Recupero, Scarti.CRecupero, PMC.Ordine," & _
                "PMC.Descrizione, PM.[Cicli/Ora], PMC.[Costo€], " & _
                "Format([Costo€]/[Cicli/Ora],'0.00') AS [CostoKG€] " & _
            "FROM (ProduzioneMacchina AS PM INNER JOIN " & _
                "ProduzioneMacchineCosto AS PMC ON PM.CdL = PMC.CdL) " & _
                "LEFT JOIN Scarti ON (PM.CdL = Scarti.CdL) AND (PM.Ldp = Scarti.Ldp) " & _
            "Where PM.Ldp = '" & Trim$(Mid$(Me.Key, 6, 4)) & "' " & _
            "ORDER BY PM.Ldp, PMC.Ordine, PM.CdL"
    Set rsCiclo = db.OpenRecordset(sqlc, dbOpenDynaset)
    If Not rsCiclo.EOF Then
      Me.Status = cExists
    Else
      Me.Status = cNotExists
    End If
  
    Do While Not rsCiclo.EOF
    
         Select Case UCase(rsCiclo("Cdl"))
            Case "C10"
              Calandra1
            Case "C20"
                Calandra2
            Case "C30"
                Calandra3
            Case "C50"
                Calandra5
            Case "B00"
                B00
            Case "BJ"
                BJ
            Case "A10" ' Accoppiatrice
                Accoppiatrice
            Case "B11" ' Taglio sleeve
                B11
            Case "E10" 'Estrusore
                Estrusore
            Case "M01" ' Metallizzatore
                Metallizzatore
            Case "Q01"  'squadratrice
                Squadratrice
            Case "S01" 'Sleeve
                Sleeve
        End Select
            
        If fA Or fS Or fM Then
            lblLS.Visible = True
            lblLR.Visible = True
            lblLT.Visible = True
            lblLSt.Visible = True
            lblLCr.Visible = True
        Else
            lblLS.Visible = False
            lblLR.Visible = False
            lblLT.Visible = False
            lblLSt.Visible = False
            lblLCr.Visible = False
        End If
    
    rsCiclo.MoveNext
    Loop
End Sub
Private Sub Macchine(Flag As Boolean)
Dim i As Integer
i = 0
                txtC1.Visible = Flag
                txtKHC1.Visible = Flag
                txtEHC1.Visible = Flag
                txtEKC1.Visible = Flag
                txtC2.Visible = Flag
                txtKHC2.Visible = Flag
                txtEHC2.Visible = Flag
                txtEKC2.Visible = Flag
                txtC3.Visible = Flag
                txtKHC3.Visible = Flag
                txtEHC3.Visible = Flag
                txtEKC3.Visible = Flag
                txtC5.Visible = Flag
                txtKHC5.Visible = Flag
                txtEHC5.Visible = Flag
                txtEKC5.Visible = Flag
                txtB00.Visible = Flag
                txtKHB00.Visible = Flag
                txtEHB00.Visible = Flag
                txtEKB00.Visible = Flag
                txtBJ.Visible = Flag
                txtKHBJ.Visible = Flag
                txtEHBJ.Visible = Flag
                txtEKBJ.Visible = Flag
                txtA10.Visible = Flag
                txtKHA10.Visible = Flag
                txtEHA10.Visible = Flag
                txtEKA10.Visible = Flag
                txtB11.Visible = Flag
                txtKHB11.Visible = Flag
                txtEHB11.Visible = Flag
                txtEKB11.Visible = Flag
                txtE10.Visible = Flag
                txtKHE10.Visible = Flag
                txtEHE10.Visible = Flag
                txtEKE10.Visible = Flag
                txtM01.Visible = Flag
                txtKHM01.Visible = Flag
                txtEHM01.Visible = Flag
                txtEKM01.Visible = Flag
                txtQ01.Visible = Flag
                txtKHQ01.Visible = Flag
                txtEHQ01.Visible = Flag
                txtEKQ01.Visible = Flag
                txtS01.Visible = Flag
                txtKHS01.Visible = Flag
                txtEHS01.Visible = Flag
                txtEKS01.Visible = Flag
'Scarti
                txtSC1.Visible = Flag
                txtSC2.Visible = Flag
                txtSC3.Visible = Flag
                txtSC5.Visible = Flag
                txtSB00.Visible = Flag
                txtSBJ.Visible = Flag
                txtSA10.Visible = Flag
                txtSB11.Visible = Flag
                txtSE10.Visible = Flag
                txtSM01.Visible = Flag
                txtSQ01.Visible = Flag
                txtSS01.Visible = Flag
'scarti costi
                txtCSC1.Visible = Flag
                txtCSC2.Visible = Flag
                txtCSC3.Visible = Flag
                txtCSC5.Visible = Flag
                txtCSE10.Visible = Flag

                txtCSSC1.Visible = Flag
                txtCSSC2.Visible = Flag
                txtCSSC3.Visible = Flag
                txtCSSC5.Visible = Flag
                txtCSSE10.Visible = Flag

'recuperi
                txtRC1.Visible = Flag
                txtCRSC1.Visible = Flag
                txtCRCC1.Visible = Flag
                txtTCC1.Visible = Flag
                txtTSC1.Visible = Flag
                
                txtRC2.Visible = Flag
                txtCRSC2.Visible = Flag
                txtCRCC2.Visible = Flag
                txtTCC2.Visible = Flag
                txtTSC2.Visible = Flag
                
                txtRC3.Visible = Flag
                txtCRSC3.Visible = Flag
                txtCRCC3.Visible = Flag
                txtTCC3.Visible = Flag
                txtTSC3.Visible = Flag
                
                txtRC5.Visible = Flag
                txtCRSC5.Visible = Flag
                txtCRCC5.Visible = Flag
                txtTCC5.Visible = Flag
                txtTSC5.Visible = Flag
                
                txtRE10.Visible = Flag
                txtCRSE10.Visible = Flag
                txtCRCE10.Visible = Flag
                txtTCE10.Visible = Flag
                txtTSE10.Visible = Flag
                
                txtRM01.Visible = Flag
                For i = 0 To txtCRSM01.Count - 1
                    txtCSM01(i).Visible = Flag
                    txtCSSM01(i).Visible = Flag
                    txtCRSM01(i).Visible = Flag
                    txtCRCM01(i).Visible = Flag
                    txtTCM01(i).Visible = Flag
                    txtTSM01(i).Visible = Flag
                Next i
                
                txtRQ01.Visible = Flag
                For i = 0 To txtCRSB00.Count - 1
                    txtCSQ01(i).Visible = Flag
                    txtCSSQ01(i).Visible = Flag
                    txtCRSQ01(i).Visible = Flag
                    txtCRCQ01(i).Visible = Flag
                    txtTCQ01(i).Visible = Flag
                    txtTSQ01(i).Visible = Flag
                Next i
                
                
                txtRS01.Visible = Flag
                For i = 0 To txtCRSS01.Count - 1
                    txtCSSS01(i).Visible = Flag
                    txtCSS01(i).Visible = Flag
                    txtCRSS01(i).Visible = Flag
                    txtCRCS01(i).Visible = Flag
                    txtTCS01(i).Visible = Flag
                    txtTSS01(i).Visible = Flag
                Next i
                
                For i = 0 To txtCRSB00.Count - 1
                    txtCSB00(i).Visible = Flag
                    txtCSSB00(i).Visible = Flag
                    txtCRSB00(i).Visible = Flag
                    txtCRCB00(i).Visible = Flag
                    txtTCB00(i).Visible = Flag
                    txtTSB00(i).Visible = Flag
                Next i
                
                txtRB00.Visible = Flag
                
                txtRB11.Visible = Flag
                For i = 0 To txtCRSB11.Count - 1
                    txtCSB11(i).Visible = Flag
                    txtCSSB11(i).Visible = Flag
                    txtCRSB11(i).Visible = Flag
                    txtCRCB11(i).Visible = Flag
                    txtTCB11(i).Visible = Flag
                    txtTSB11(i).Visible = Flag
                Next i
                
                txtRBJ.Visible = Flag
           For i = 0 To txtCRSBJ.Count - 1
                    txtCSBJ(i).Visible = Flag
                    txtCSSBJ(i).Visible = Flag
                    txtCRSBJ(i).Visible = Flag
                    txtCRCBJ(i).Visible = Flag
                    txtTCBJ(i).Visible = Flag
                    txtTSBJ(i).Visible = Flag
                Next i
                
                txtRA10.Visible = Flag
                For i = 0 To txtCRSA10.Count - 1
                    txtCSSA10(i).Visible = Flag
                    txtCSA10(i).Visible = Flag
                    txtCRSA10(i).Visible = Flag
                    txtCRCA10(i).Visible = Flag
                    txtTCA10(i).Visible = Flag
                    txtTSA10(i).Visible = Flag
                Next i
                txtRCC1.Visible = Flag
                txtRCC2.Visible = Flag
                txtRCC3.Visible = Flag
                txtRCC5.Visible = Flag
                txtRCE10.Visible = Flag
                txtRCS01.Visible = Flag
                txtRCA10.Visible = Flag
                txtRCM01.Visible = Flag
                txtRCB00.Visible = Flag
                txtRCBJ.Visible = Flag
                txtRCQ01.Visible = Flag
                txtRCB11.Visible = Flag
                
                
 
 
 ' shape
                shpC2.Visible = Flag
                shpC3.Visible = Flag
                shpC5.Visible = Flag
                
'soltanto per gli accoppiati
                txtMP.Visible = Flag
                txtMPD.Visible = Flag
                txtMPC.Visible = Flag
                txtMPS.Visible = Flag
                txtMPPS.Visible = Flag
                txtSP.Visible = Flag


End Sub
Private Function Posizione(tE As TextBox, fV As Boolean, iL As Double, iT As Double)

With tE
    .Visible = fV
    .Left = iL
    .Top = iT
End With

End Function
Private Sub Calandra1()
                txtC1.Visible = True
                txtKHC1.Visible = True
                txtEHC1.Visible = True
                txtEKC1.Visible = True
                txtSC1.Visible = True
                txtCSC1.Visible = True
                txtCSSC1.Visible = True
                txtRC1.Visible = True
                txtCRSC1.Visible = True
                txtCRCC1.Visible = True
                txtRCC1.Visible = True
                If Not fRicalcola Then
                    txtC1.Text = rsCiclo("Descrizione")
                    txtKHC1.Text = rsCiclo("Cicli/Ora")
                    txtEHC1.Text = rsCiclo("Costo€")
                    txtEKC1.Text = rsCiclo("CostoKG€")
'                    txtSC1.Text = rsCiclo("Scarto")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"

                Else
                    txtC1.Text = txtC1.Text
                    txtKHC1.Text = txtKHC1.Text
                    txtEHC1.Text = txtEHC1.Text
                    txtEKC1.Text = txtEHC1.Text / txtKHC1
                    txtSC1.Text = txtSC1.Text
                    sScarto = "0"
                    sRecupero = "0"

                End If
                
                
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSC1.Text <> sScarto Then
                    dPrezzoStd = CDbl(txtMescola(7).Text) ' + CDbl(txtEKC1.Text)
                    
                    If Not fRicalcola Then
                        dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC1.Text)
                        txtSC1.Text = rsCiclo("Scarto")
                    Else
                        If Not IsNull(txtCCMa514.Text) And txtCCMa514.Text <> "" Then
                            dPrezzoCrr = CDbl(txtMescola(8).Text) + CalcolaMa514
                        Else
                            dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC1.Text)
                        End If
                        txtSC1.Text = txtSC1.Text
                    
                    
                    
                    End If
                    
                    txtCSC1.Text = CalcolaScarto(txtSC1.Text, dPrezzoCrr)
                    txtCSSC1.Text = CalcolaScarto(txtSC1.Text, dPrezzoStd)
                Else
                    txtSC1.Text = "0"
                    txtCSSC1.Text = "0"  'txtMescola(7).Text
                    txtCSC1.Text = "0"  'txtMescola(8).Text
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRC1.Text <> sRecupero Then
                    If Not fRicalcola Then
                      txtRC1.Text = rsCiclo("Recupero")
                      txtRCC1.Text = rsCiclo("CRecupero")
                    Else
                        txtRC1.Text = txtRC1.Text
                        txtRCC1.Text = txtRCC1.Text
                    End If
                    
                    txtCRSC1.Text = CDbl((txtSC1.Text) / 100) * CDbl((txtRC1.Text / 100)) * CDbl(txtRCC1.Text)
                    txtCRCC1.Text = CDbl((txtSC1.Text) / 100) * CDbl((txtRC1.Text / 100)) * CDbl(txtRCC1.Text)
                Else
                    txtRC1.Text = "0"
                    txtCRSC1.Text = "0"
                    txtCRCC1.Text = "0"
                End If
'   Totali
                txtTCC1.Visible = True
                txtTSC1.Visible = True

                txtTCC1.Text = CDbl(txtCSC1.Text) - CDbl(txtCRCC1.Text) + dPrezzoCrr + CDbl(txtEKC1.Text)
                txtTSC1.Text = CDbl(txtCSSC1.Text) - CDbl(txtCRSC1.Text) + dPrezzoStd + CDbl(txtEKC1.Text)
'attivo i flag
                f1 = True
End Sub
Private Sub Calandra2()
                txtC2.Visible = True
                txtKHC2.Visible = True
                txtEHC2.Visible = True
                txtEKC2.Visible = True
                txtSC2.Visible = True
                txtCSC2.Visible = True
                txtCSSC2.Visible = True
                txtRC2.Visible = True
                txtCRSC2.Visible = True
                txtCRCC2.Visible = True
                txtRCC2.Visible = True
                txtC2.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                    txtKHC2.Text = rsCiclo("Cicli/Ora")
                    txtEHC2.Text = rsCiclo("Costo€")
                    txtEKC2.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"
                    
                Else
                    txtKHC2.Text = txtKHC2.Text
                    txtEHC2.Text = txtEHC2.Text
                    txtEKC2.Text = txtEHC2.Text / txtKHC2.Text
                    sScarto = "0"
                    sRecupero = "0"
                
                End If
                
                            
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSC2.Text <> sScarto Then

                    dPrezzoStd = CDbl(txtMescola(7).Text) ' + CDbl(txtEKC2.Text)
                    If Not fRicalcola Then
                        dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC2.Text)
                        txtSC2.Text = rsCiclo("Scarto")
                    Else
                        If Not IsNull(txtCCMa514.Text) And txtCCMa514.Text <> "" Then
                            dPrezzoCrr = CDbl(txtMescola(8).Text) + CalcolaMa514
                        Else
                            dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC1.Text)
                        End If
    
                        
                        txtSC2.Text = txtSC2.Text
                    End If
                    txtCSC2.Text = CalcolaScarto(txtSC2.Text, dPrezzoCrr)
                    txtCSSC2.Text = CalcolaScarto(txtSC2.Text, dPrezzoStd)
                Else
                    txtSC2.Text = "0"
                    txtCSSC2.Text = "0" 'txtMescola(7).Text
                    txtCSC2.Text = "0"  'txtMescola(8).Text
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRC2.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRC2.Text = rsCiclo("Recupero")
                        txtRCC2.Text = rsCiclo("CRecupero")
                    
                    Else
                        txtRC2.Text = txtRC2.Text
                        txtRCC2.Text = txtRCC2.Text
                    End If
'                    txtRC2.Text = rsCiclo("Recupero")
                    txtCRSC2.Text = CDbl((txtSC2.Text) / 100) * CDbl((txtRC2.Text / 100)) * CDbl(txtRCC2.Text)
                    txtCRCC2.Text = CDbl((txtSC2.Text) / 100) * CDbl((txtRC2.Text / 100)) * CDbl(txtRCC2.Text)
                Else
                    txtRC2.Text = "0"
                    txtCRSC2.Text = "0"
                    txtCRCC2.Text = "0"
                End If
'   Totali
                txtTCC2.Visible = True
                txtTSC2.Visible = True
'                If fRicalcola Then
'                    dPrezzoCrr = CDbl(txtMescola(8).Text)
                    dPrezzoStd = CDbl(txtMescola(7).Text)
'                End If
                
                    
                    txtTCC2.Text = CDbl(txtCSC2.Text) - CDbl(txtCRCC2.Text) + dPrezzoCrr + CDbl(txtEKC2.Text)
                    txtTSC2.Text = CDbl(txtCSSC2.Text) - CDbl(txtCRSC2.Text) + dPrezzoStd + CDbl(txtEKC2.Text)
                    


'attivo i flag
                f2 = True
End Sub
Private Sub Calandra3()
                txtC3.Visible = True
                txtKHC3.Visible = True
                txtEHC3.Visible = True
                txtEKC3.Visible = True
                txtSC3.Visible = True
                txtCSC3.Visible = True
                txtCSSC3.Visible = True
                txtRC3.Visible = True
                txtCRSC3.Visible = True
                txtCRCC3.Visible = True
                txtRCC3.Visible = True
                txtC3.Text = rsCiclo("Descrizione")
                    
                If Not fRicalcola Then
                
                    txtKHC3.Text = rsCiclo("Cicli/Ora")
                    txtEHC3.Text = rsCiclo("Costo€")
                    txtEKC3.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"

                Else
                    txtKHC3.Text = txtKHC3.Text
                    txtEHC3.Text = txtEHC3.Text
                    txtEKC3.Text = txtEHC3.Text / txtKHC3.Text
                    sScarto = "0"
                    sRecupero = "0"
                
                End If
                
                
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSC3.Text <> sScarto Then
 
                    dPrezzoStd = CDbl(txtMescola(7).Text) ' + CDbl(txtEKC3.Text)
                    If Not fRicalcola Then
                        dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC3.Text)
                        txtSC3.Text = rsCiclo("Scarto")
                    
                    Else
                        If Not IsNull(txtCCMa514.Text) And txtCCMa514.Text <> "" Then
                            dPrezzoCrr = CDbl(txtMescola(8).Text) + CalcolaMa514
                        Else
                            dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC1.Text)
                        End If
                        
                        txtSC3.Text = txtSC3.Text
                    End If
                    
                    txtCSC3.Text = CalcolaScarto(txtSC3.Text, dPrezzoCrr)
                    txtCSSC3.Text = CalcolaScarto(txtSC3.Text, dPrezzoStd)
                Else
                    txtSC3.Text = "0"
                    txtCSSC3.Text = "0" ' txtMescola(7).Text
                    txtCSC3.Text = "0"  'txtMescola(8).Text
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRC3.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRC3.Text = rsCiclo("Recupero")
                        txtRCC3.Text = rsCiclo("CRecupero")

                    Else
                        txtRC3.Text = txtRC3.Text
                        txtRCC3.Text = txtRCC3.Text

                    End If
                    txtCRSC3.Text = CDbl((txtSC3.Text) / 100) * CDbl((txtRC3.Text / 100)) * CDbl(txtRCC3.Text)
                    txtCRCC3.Text = CDbl((txtSC3.Text) / 100) * CDbl((txtRC3.Text / 100)) * CDbl(txtRCC3.Text)
                Else
                    txtRC3.Text = "0"
                    txtCRSC3.Text = "0"
                    txtCRCC3.Text = "0"
                    
                End If
'   Totali
                txtTCC3.Visible = True
                txtTSC3.Visible = True

                txtTCC3.Text = CDbl(txtCSC3.Text) - CDbl(txtCRCC3.Text) + dPrezzoCrr + CDbl(txtEKC3.Text)
                txtTSC3.Text = CDbl(txtCSSC3.Text) - CDbl(txtCRSC3.Text) + dPrezzoStd + CDbl(txtEKC3.Text)
'attivo i flag
                f3 = True
                
End Sub

Private Sub Calandra5()
                txtC5.Visible = True
                txtKHC5.Visible = True
                txtEHC5.Visible = True
                txtEKC5.Visible = True
                txtSC5.Visible = True
                txtCSC5.Visible = True
                txtCSSC5.Visible = True
                txtRC5.Visible = True
                txtCRSC5.Visible = True
                txtCRCC5.Visible = True
                txtRCC5.Visible = True
                txtC5.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                
                    txtKHC5.Text = rsCiclo("Cicli/Ora")
                    txtEHC5.Text = rsCiclo("Costo€")
                    txtEKC5.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"

                Else
                    txtKHC5.Text = txtKHC5.Text
                    txtEHC5.Text = txtEHC5.Text
                    txtEKC5.Text = txtEHC5.Text / txtKHC5.Text
                    sScarto = "0"
                    sRecupero = "0"

                End If
                
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSC5.Text <> sScarto Then
'                    dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC5.Text)
                    dPrezzoStd = CDbl(txtMescola(7).Text) ' + CDbl(txtEKC5.Text)
                    If Not fRicalcola Then
                        dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKE10.Text)
                        txtSC5.Text = rsCiclo("Scarto")
                    Else
                        If Not IsNull(txtCCMa514.Text) And txtCCMa514.Text <> "" Then
                            dPrezzoCrr = CDbl(txtMescola(8).Text) + CalcolaMa514
                        Else
                            dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC1.Text)
                        End If
                        
                        txtSC5.Text = txtSC5.Text
                    End If
                    
                    txtCSC5.Text = CalcolaScarto(txtSC5.Text, dPrezzoCrr)
                    txtCSSC5.Text = CalcolaScarto(txtSC5.Text, dPrezzoStd)
                Else
                    txtSC5.Text = "0"
                    txtCSSC5.Text = "0" ' txtMescola(7).Text
                    txtCSC5.Text = "0"  'txtMescola(8).Text
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRC5.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRC5.Text = rsCiclo("Recupero")
                        txtRCC5.Text = rsCiclo("CRecupero")
                    Else
                        txtRC5.Text = txtRC5.Text
                        txtRCC5.Text = txtRCC5.Text
                    End If
                    txtCRSC5.Text = CDbl((txtSC5.Text) / 100) * CDbl((txtRC5.Text / 100)) * CDbl(txtRCC5.Text)
                    txtCRCC5.Text = CDbl((txtSC5.Text) / 100) * CDbl((txtRC5.Text / 100)) * CDbl(txtRCC5.Text)
                Else
                    txtRC5.Text = "0"
                    txtCRSC5.Text = "0"
                    txtCRCC5.Text = "0"
                End If
'   Totali
                txtTCC5.Visible = True
                txtTSC5.Visible = True

                txtTCC5.Text = CDbl(txtCSC5.Text) - CDbl(txtCRCC5.Text) + dPrezzoCrr + CDbl(txtEKC5.Text)
                txtTSC5.Text = CDbl(txtCSSC5.Text) - CDbl(txtCRSC5.Text) + dPrezzoStd + CDbl(txtEKC5.Text)
'attivo i flag
                f5 = True
End Sub
Private Sub Estrusore()
'                posizione
                txtE10.Top = txtC1.Top
                txtKHE10.Top = txtKHC1.Top
                txtEHE10.Top = txtEHC1.Top
                txtEKE10.Top = txtEKC1.Top
                txtSE10.Top = txtSC1.Top
                txtCSE10.Top = txtCSC1.Top
                txtCSSE10.Top = txtCSSC1.Top
                txtRE10.Top = txtRC1.Top
                txtRCE10.Top = txtRCC1.Top
                txtCRSE10.Top = txtCRSC1.Top
                txtCRCE10.Top = txtCRCC1.Top
                txtTCE10.Top = txtTCC1.Top
                txtTSE10.Top = txtTSC1.Top
                
                txtE10.Left = txtC1.Left
                txtKHE10.Left = txtKHC1.Left
                txtEHE10.Left = txtEHC1.Left
                txtEKE10.Left = txtEKC1.Left
                txtSE10.Left = txtSC1.Left
                txtCSE10.Left = txtCSC1.Left
                txtCSSE10.Left = txtCSSC1.Left
                txtRE10.Left = txtRC1.Left
                txtRCE10.Left = txtRCC1.Left
                txtCRSE10.Left = txtCRSC1.Left
                txtCRCE10.Left = txtCRCC1.Left
                txtTCE10.Left = txtTCC1.Left
                txtTSE10.Left = txtTSC1.Left
                
                txtE10.Visible = True
                txtKHE10.Visible = True
                txtEHE10.Visible = True
                txtEKE10.Visible = True
                txtSE10.Visible = True
                txtCSE10.Visible = True
                txtCSSE10.Visible = True
                txtRE10.Visible = True
                txtCRSE10.Visible = True
                txtCRCE10.Visible = True
                txtRCE10.Visible = True
             
                txtE10.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                
                    txtKHE10.Text = rsCiclo("Cicli/Ora")
                    txtEHE10.Text = rsCiclo("Costo€")
                    txtEKE10.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"
               
                Else
                    txtKHE10.Text = txtKHE10.Text
                    txtEHE10.Text = txtEHE10.Text
                    txtEKE10.Text = txtEHE10.Text / txtKHE10.Text
                    sScarto = "0"
                    sRecupero = "0"
               
                End If
                
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSE10.Text <> sScarto Then
                    dPrezzoStd = CDbl(txtMescola(7).Text) ' + CDbl(txtEKE10.Text)
                    If Not fRicalcola Then
                        dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKE10.Text)
                        txtSE10.Text = rsCiclo("Scarto")
                    Else
                        If Not IsNull(txtCCMa514.Text) And txtCCMa514.Text <> "" Then
                            dPrezzoCrr = CDbl(txtMescola(8).Text) + CalcolaMa514
                        Else
                            dPrezzoCrr = CDbl(txtMescola(8).Text) ' + CDbl(txtEKC1.Text)
                        End If
                       
                        txtSE10.Text = txtSE10.Text
                    End If
                    
                    txtCSE10.Text = CalcolaScarto(txtSE10.Text, dPrezzoCrr)
                    txtCSSE10.Text = CalcolaScarto(txtSE10.Text, dPrezzoStd)
                Else
                    txtSE10.Text = "0"
                    txtCSSE10.Text = "0" 'txtMescola(7).Text
                    txtCSE10.Text = "0" ' txtMescola(8).Text
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRE10.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRE10.Text = rsCiclo("Recupero")
                        txtRCE10.Text = rsCiclo("CRecupero")
                    Else
                        txtRE10.Text = txtRE10.Text
                        txtRCE10.Text = txtRCE10.Text
                    End If
                    
                    txtCRSE10.Text = CDbl((txtSE10.Text) / 100) * CDbl((txtRE10.Text / 100)) * CDbl(txtRCE10.Text)
                    txtCRCE10.Text = CDbl((txtSE10.Text) / 100) * CDbl((txtRE10.Text / 100)) * CDbl(txtRCE10.Text)
                Else
                    txtRE10.Text = "0"
                    txtCRSE10.Text = "0"
                    txtCRCE10.Text = "0"
                End If
'   Totali
                txtTCE10.Visible = True
                txtTSE10.Visible = True

                txtTCE10.Text = CDbl(txtCSE10.Text) - CDbl(txtCRCE10.Text) + dPrezzoCrr + CDbl(txtEKE10.Text)
                txtTSE10.Text = CDbl(txtCSSE10.Text) - CDbl(txtCRSE10.Text) + dPrezzoStd + CDbl(txtEKE10.Text)
'attivo i flag
                fE = True
                iLinea = 0
End Sub
Private Sub Metallizzatore()
'attivo il flag per le tagliatrici
                fM = True


'                posizione
                txtM01.Top = txtS01.Top
                txtKHM01.Top = txtKHS01.Top
                txtEHM01.Top = txtEHS01.Top
                txtEKM01.Top = txtEKS01.Top
                txtSM01.Top = txtSS01.Top
                txtCSM01(0).Top = txtCSS01(0).Top
                txtCSSM01(0).Top = txtCSSS01(0).Top
                txtRM01.Top = txtRS01.Top
                txtRCM01.Top = txtRCS01.Top
                txtCRSM01(0).Top = txtCRSS01(0).Top
                txtCRCM01(0).Top = txtCRCS01(0).Top
                txtTCM01(0).Top = txtTCS01(0).Top
                txtTSM01(0).Top = txtTSS01(0).Top
               
                
                txtM01.Left = txtS01.Left
                txtKHM01.Left = txtKHS01.Left
                txtEHM01.Left = txtEHS01.Left
                txtEKM01.Left = txtEKS01.Left
                txtSM01.Left = txtSS01.Left
                txtCSM01(0).Left = txtCSS01(0).Left
                txtCSSM01(0).Left = txtCSSS01(0).Left
                txtRM01.Left = txtRS01.Left
                txtRCM01.Left = txtRCS01.Left
                txtCRSM01(0).Left = txtCRSS01(0).Left
                txtCRCM01(0).Left = txtCRCS01(0).Left
                txtTCM01(0).Left = txtTCS01(0).Left
                txtTSM01(0).Left = txtTSS01(0).Left
                
                
                
                txtM01.Visible = True
                txtKHM01.Visible = True
                txtEHM01.Visible = True
                txtEKM01.Visible = True
                txtSM01.Visible = True
                txtCSM01(0).Visible = True
                txtCSSM01(0).Visible = True
                txtRM01.Visible = True
                txtRCM01.Visible = True
                txtCRSM01(0).Visible = True
                txtCRCM01(0).Visible = True
                
                txtM01.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                    txtKHM01.Text = rsCiclo("Cicli/Ora")
                    txtEHM01.Text = rsCiclo("Costo€")
                    txtEKM01.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"

                Else
                    txtKHM01.Text = txtKHM01.Text
                    txtEHM01.Text = txtEHM01.Text
                    txtEKM01.Text = txtEHM01.Text / txtKHM01.Text
                    sScarto = "0"
                    sRecupero = "0"
               
                End If
                
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSM01.Text <> sScarto Then
                    If Not fRicalcola Then
                        txtSM01.Text = rsCiclo("Scarto")
                    Else
                        txtSM01.Text = txtSM01.Text
                    End If
                    
                    If f1 Then
                        dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKM01.Text)
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC1.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                        End If
                        
                        txtCSM01(0).Text = CalcolaScarto(txtSM01.Text, dPrezzoCrr)
                        txtCSSM01(0).Text = CalcolaScarto(txtSM01.Text, dPrezzoStd)
                    End If
                    If f2 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC2.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKM01.Text)
                        txtCSM01(1).Text = CalcolaScarto(txtSM01.Text, dPrezzoCrr)
                        txtCSSM01(1).Text = CalcolaScarto(txtSM01.Text, dPrezzoStd)
                    End If
                    If f3 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC3.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKM01.Text)
                        txtCSM01(2).Text = CalcolaScarto(txtSM01.Text, dPrezzoCrr)
                        txtCSSM01(2).Text = CalcolaScarto(txtSM01.Text, dPrezzoStd)
                    End If
                    If f5 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC5.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                        End If

                        dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKM01.Text)
                        txtCSM01(3).Text = CalcolaScarto(txtSM01.Text, dPrezzoCrr)
                        txtCSSM01(3).Text = CalcolaScarto(txtSM01.Text, dPrezzoStd)
                    End If
                    If fE Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCE10.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                        End If

                        dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKM01.Text)
                        txtCSM01(0).Text = CalcolaScarto(txtSM01.Text, dPrezzoCrr)
                        txtCSSM01(0).Text = CalcolaScarto(txtSM01.Text, dPrezzoStd)
                    End If
               
                Else
                    
                    txtSM01.Text = "0"
                    For i = 0 To txtCSSM01.Count - 1
                        txtCSSM01(i).Text = "0"
                        txtCSM01(i).Text = "0"
                    Next i
                End If
                   
                    
                 If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRM01.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRM01.Text = rsCiclo("Recupero")
                        txtRCM01.Text = rsCiclo("CRecupero")

                    Else
                        txtRM01.Text = txtRM01.Text
                        txtRCM01.Text = txtRCM01.Text
                    End If
                For i = 0 To txtCRSM01.Count - 1
                    txtCRSM01(i).Text = CDbl((txtSM01.Text) / 100) * CDbl((txtRM01.Text / 100)) * CDbl(txtRCM01.Text)
                    txtCRCM01(i).Text = CDbl((txtSM01.Text) / 100) * CDbl((txtRM01.Text / 100)) * CDbl(txtRCM01.Text)
                Next i
                Else
                    txtRM01.Text = "0"
                    For i = 0 To txtCRSM01.Count - 1
                        txtCRSM01(i).Text = "0"
                        txtCRCM01(i).Text = "0"
                    Next i
                End If
'   Totali
                
                
                If f1 Then
'                    dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKM01.Text)
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC1.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                    End If

                    dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKM01.Text)
                    txtTCM01(0).Text = CDbl(txtCSM01(0).Text) - CDbl(txtCRCM01(0).Text) + dPrezzoCrr + CDbl(txtEKM01.Text)
                    txtTSM01(0).Text = CDbl(txtCSSM01(0).Text) - CDbl(txtCRSM01(0).Text) + dPrezzoStd + CDbl(txtEKM01.Text)
                    txtCRSM01(0).Visible = True
                    txtCRCM01(0).Visible = True
                    txtCSM01(0).Visible = True
                    txtCSSM01(0).Visible = True
                    txtTCM01(0).Visible = True
                    txtTSM01(0).Visible = True
                End If
                If f2 Then
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC2.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                    End If

'                    dPrezzoCrr = CDbl(txtTCC2.Text) ' + CDbl(txtEKM01.Text)
                    dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKM01.Text)
                    txtTCM01(1).Text = CDbl(txtCSM01(1).Text) - CDbl(txtCRCM01(1).Text) + dPrezzoCrr + CDbl(txtEKM01.Text)
                    txtTSM01(1).Text = CDbl(txtCSSM01(1).Text) - CDbl(txtCRSM01(1).Text) + dPrezzoStd + CDbl(txtEKM01.Text)
                    Posizione txtCRSM01(1), True, txtCRSC2.Left, txtCRSM01(0).Top
                    Posizione txtCRCM01(1), True, txtCRCC2.Left, txtCRCM01(0).Top
                    Posizione txtCSM01(1), True, txtCSC2.Left, txtCSM01(0).Top
                    Posizione txtCSSM01(1), True, txtCSSC2.Left, txtCSSM01(0).Top
                    Posizione txtTCM01(1), True, txtTCC2.Left, txtTCM01(0).Top
                    Posizione txtTSM01(1), True, txtTSC2.Left, txtTSM01(0).Top
                    shpC2.Visible = True
                End If
                If f3 Then
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC3.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                    End If
                   
'                    dPrezzoCrr = CDbl(txtTCC3.Text) ' + CDbl(txtEKM01.Text)
                    dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKM01.Text)
                    txtTCM01(2).Text = CDbl(txtCSM01(2).Text) - CDbl(txtCRCM01(2).Text) + dPrezzoCrr + CDbl(txtEKM01.Text)
                    txtTSM01(2).Text = CDbl(txtCSSM01(2).Text) - CDbl(txtCRSM01(2).Text) + dPrezzoStd + CDbl(txtEKM01.Text)
                    Posizione txtCRSM01(2), True, txtCRSC3.Left, txtCRSM01(0).Top
                    Posizione txtCRCM01(2), True, txtCRCC3.Left, txtCRCM01(0).Top
                    Posizione txtCSM01(2), True, txtCSC3.Left, txtCSM01(0).Top
                    Posizione txtCSSM01(2), True, txtCSSC3.Left, txtCSSM01(0).Top
                    Posizione txtTCM01(2), True, txtTCC3.Left, txtTCM01(0).Top
                    Posizione txtTSM01(2), True, txtTSC3.Left, txtTSM01(0).Top
                    shpC3.Visible = True
                End If
                If f5 Then
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC5.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                    End If
                   
'                    dPrezzoCrr = CDbl(txtTCC5.Text) ' + CDbl(txtEKM01.Text)
                    dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKM01.Text)
                    txtTCM01(3).Text = CDbl(txtCSM01(3).Text) - CDbl(txtCRCM01(3).Text) + dPrezzoCrr + CDbl(txtEKM01.Text)
                    txtTSM01(3).Text = CDbl(txtCSSM01(3).Text) - CDbl(txtCRSM01(3).Text) + dPrezzoStd + CDbl(txtEKM01.Text)
                    Posizione txtCRSM01(3), True, txtCRSC5.Left, txtCRSM01(0).Top
                    Posizione txtCRCM01(3), True, txtCRCC5.Left, txtCRCM01(0).Top
                    Posizione txtCSM01(3), True, txtCSC5.Left, txtCSM01(0).Top
                    Posizione txtCSSM01(3), True, txtCSSC5.Left, txtCSSM01(0).Top
                    Posizione txtTCM01(3), True, txtTCC5.Left, txtTCM01(0).Top
                    Posizione txtTSM01(3), True, txtTSC5.Left, txtTSM01(0).Top
                    shpC5.Visible = True
                End If
                If fE Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCE10.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                        End If
                   
'                    dPrezzoCrr = CDbl(txtTCE10.Text) ' + CDbl(txtEKM01.Text)
                    dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKM01.Text)
                    txtTCM01(0).Text = CDbl(txtCSM01(0).Text) - CDbl(txtCRCM01(0).Text) + dPrezzoCrr + CDbl(txtEKM01.Text)
                    txtTSM01(0).Text = CDbl(txtCSSM01(0).Text) - CDbl(txtCRSM01(0).Text) + dPrezzoStd + CDbl(txtEKM01.Text)
                    txtCRSM01(0).Visible = True
                    txtCRCM01(0).Visible = True
                    txtCSM01(0).Visible = True
                    txtCSSM01(0).Visible = True
                    txtTCM01(0).Visible = True
                    txtTSM01(0).Visible = True
'                    txtTCM01.Count = 1
                End If
                
End Sub

Private Sub Accoppiatrice()
'attivo il flag per le tagliatrici
                fA = True

'                posizione
                               
                txtA10.Top = txtS01.Top
                txtKHA10.Top = txtKHS01.Top
                txtEHA10.Top = txtEHS01.Top
                txtEKA10.Top = txtEKS01.Top
                txtSA10.Top = txtSS01.Top
                txtCSA10(0).Top = txtCSS01(0).Top
                txtCSSA10(0).Top = txtCSSS01(0).Top
                txtRA10.Top = txtRS01.Top
                txtRCA10.Top = txtRCS01.Top
                txtCRSA10(0).Top = txtCRSS01(0).Top
                txtCRCA10(0).Top = txtCRCS01(0).Top
                txtTCA10(0).Top = txtTCS01(0).Top
                txtTSA10(0).Top = txtTSS01(0).Top
            
                txtA10.Left = txtS01.Left
                txtKHA10.Left = txtKHS01.Left
                txtEHA10.Left = txtEHS01.Left
                txtEKA10.Left = txtEKS01.Left
                txtSA10.Left = txtSS01.Left
                txtCSA10(0).Left = txtCSS01(0).Left
                txtCSSA10(0).Left = txtCSSS01(0).Left
                txtRA10.Left = txtRS01.Left
                txtRCA10.Left = txtRCS01.Left
                txtCRSA10(0).Left = txtCRSS01(0).Left
                txtCRCA10(0).Left = txtCRCS01(0).Left
                txtTCA10(0).Left = txtTCS01(0).Left
                txtTSA10(0).Left = txtTSS01(0).Left
              
                
                txtA10.Visible = True
                txtKHA10.Visible = True
                txtEHA10.Visible = True
                txtEKA10.Visible = True
                txtSA10.Visible = True
                txtCSA10(0).Visible = True
                txtCSSA10(0).Visible = True
                txtRA10.Visible = True
                txtRCA10.Visible = True
                txtCRSA10(0).Visible = True
                txtCRCA10(0).Visible = True
               
                txtA10.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                    txtKHA10.Text = rsCiclo("Cicli/Ora")
                    txtEHA10.Text = rsCiclo("Costo€")
                    txtEKA10.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"

                Else
                    txtKHA10.Text = txtKHA10.Text
                    txtEHA10.Text = txtEHA10.Text
                    txtEKA10.Text = txtEHA10.Text / txtKHA10.Text
                    sScarto = "0"
                    sRecupero = "0"

                End If
                
               If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSA10.Text <> sScarto Then
                    If Not fRicalcola Then
                        txtSA10.Text = rsCiclo("Scarto")
                    Else
                        txtSA10.Text = txtSA10.Text
                    End If
                    If f1 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC1.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                        End If
'                       dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKA10.Text)
                        dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKA10.Text)
                        txtCSA10(0).Text = CalcolaScarto(txtSA10.Text, dPrezzoCrr)
                        txtCSSA10(0).Text = CalcolaScarto(txtSA10.Text, dPrezzoStd)
                    End If
                    If f2 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC2.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKA10.Text)
                        txtCSA10(1).Text = CalcolaScarto(txtSA10.Text, dPrezzoCrr)
                        txtCSSA10(1).Text = CalcolaScarto(txtSA10.Text, dPrezzoStd)
                    End If
                    If f3 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC3.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKA10.Text)
                        txtCSA10(2).Text = CalcolaScarto(txtSA10.Text, dPrezzoCrr)
                        txtCSSA10(2).Text = CalcolaScarto(txtSA10.Text, dPrezzoStd)
                    End If
                    If f5 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC5.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKA10.Text)
                        txtCSA10(3).Text = CalcolaScarto(txtSA10.Text, dPrezzoCrr)
                        txtCSSA10(3).Text = CalcolaScarto(txtSA10.Text, dPrezzoStd)
                    End If
                    If fE Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCE10.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKA10.Text)
                        txtCSA10(0).Text = CalcolaScarto(txtSA10.Text, dPrezzoCrr)
                        txtCSSA10(0).Text = CalcolaScarto(txtSA10.Text, dPrezzoStd)
                    End If
                
                
                Else
                    
                    txtSA10.Text = "0"
                    For i = 0 To txtCSSA10.Count - 1
                        txtCSSA10(i).Text = "0"
                        txtCSA10(i).Text = "0"
                    Next i
                End If
                   
                    
                 If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRA10.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRA10.Text = rsCiclo("Recupero")
                        txtRCA10.Text = rsCiclo("CRecupero")

                    Else
                        txtRA10.Text = txtRA10.Text
                        txtRCA10.Text = txtRCA10.Text
                    End If
                For i = 0 To txtCRSA10.Count - 1
                    txtCRSA10(i).Text = CDbl((txtSA10.Text) / 100) * CDbl((txtRA10.Text / 100)) * CDbl(txtRCA10.Text)
                    txtCRCA10(i).Text = CDbl((txtSA10.Text) / 100) * CDbl((txtRA10.Text / 100)) * CDbl(txtRCA10.Text)
                Next i
                Else
                    txtRA10.Text = "0"
                    For i = 0 To txtCRSA10.Count - 1
                        txtCRSA10(i).Text = "0"
                        txtCRCA10(i).Text = "0"
                    Next i
                End If
'   Totali
                
                
                If f1 Then
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC1.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                    End If
                    dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKA10.Text)
                    txtTCA10(0).Text = CDbl(txtCSA10(0).Text) - CDbl(txtCRCA10(0).Text) + dPrezzoCrr + CDbl(txtEKA10.Text)
                    txtTSA10(0).Text = CDbl(txtCSSA10(0).Text) - CDbl(txtCRSA10(0).Text) + dPrezzoStd + CDbl(txtEKA10.Text)
                    txtCRSA10(0).Visible = True
                    txtCRCA10(0).Visible = True
                    txtCSA10(0).Visible = True
                    txtCSSA10(0).Visible = True
                    txtTCA10(0).Visible = True
                    txtTSA10(0).Visible = True
                End If
                If f2 Then
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC2.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                    End If
                    dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKA10.Text)
                    txtTCA10(1).Text = CDbl(txtCSA10(1).Text) - CDbl(txtCRCA10(1).Text) + dPrezzoCrr + CDbl(txtEKA10.Text)
                    txtTSA10(1).Text = CDbl(txtCSSA10(1).Text) - CDbl(txtCRSA10(1).Text) + dPrezzoStd + CDbl(txtEKA10.Text)
                    Posizione txtCRSA10(1), True, txtCRSC2.Left, txtCRSA10(0).Top
                    Posizione txtCRCA10(1), True, txtCRCC2.Left, txtCRCA10(0).Top
                    Posizione txtCSA10(1), True, txtCSC2.Left, txtCSA10(0).Top
                    Posizione txtCSSA10(1), True, txtCSSC2.Left, txtCSSA10(0).Top
                    Posizione txtTCA10(1), True, txtTCC2.Left, txtTCA10(0).Top
                    Posizione txtTSA10(1), True, txtTSC2.Left, txtTSA10(0).Top
                    shpC2.Visible = True
                End If
                If f3 Then
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC3.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                    End If
                    dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKA10.Text)
                    txtTCA10(2).Text = CDbl(txtCSA10(2).Text) - CDbl(txtCRCA10(2).Text) + dPrezzoCrr + CDbl(txtEKA10.Text)
                    txtTSA10(2).Text = CDbl(txtCSSA10(2).Text) - CDbl(txtCRSA10(2).Text) + dPrezzoStd + CDbl(txtEKA10.Text)
                    Posizione txtCRSA10(2), True, txtCRSC3.Left, txtCRSA10(0).Top
                    Posizione txtCRCA10(2), True, txtCRCC3.Left, txtCRCA10(0).Top
                    Posizione txtCSA10(2), True, txtCSC3.Left, txtCSA10(0).Top
                    Posizione txtCSSA10(2), True, txtCSSC3.Left, txtCSSA10(0).Top
                    Posizione txtTCA10(2), True, txtTCC3.Left, txtTCA10(0).Top
                    Posizione txtTSA10(2), True, txtTSC3.Left, txtTSA10(0).Top
                    shpC3.Visible = True
                End If
                If f5 Then
                    If Not fAcc Then    ' Controllo se accoppiato
                        dPrezzoCrr = CDbl(txtTCC5.Text)
                    Else
                        dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                    End If
                    dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKA10.Text)
                    txtTCA10(3).Text = CDbl(txtCSA10(3).Text) - CDbl(txtCRCA10(3).Text) + dPrezzoCrr + CDbl(txtEKA10.Text)
                    txtTSA10(3).Text = CDbl(txtCSSA10(3).Text) - CDbl(txtCRSA10(3).Text) + dPrezzoStd + CDbl(txtEKA10.Text)
                    Posizione txtCRSA10(3), True, txtCRSC5.Left, txtCRSA10(0).Top
                    Posizione txtCRCA10(3), True, txtCRCC5.Left, txtCRCA10(0).Top
                    Posizione txtCSA10(3), True, txtCSC5.Left, txtCSA10(0).Top
                    Posizione txtCSSA10(3), True, txtCSSC5.Left, txtCSSA10(0).Top
                    Posizione txtTCA10(3), True, txtTCC5.Left, txtTCA10(0).Top
                    Posizione txtTSA10(3), True, txtTSC5.Left, txtTSA10(0).Top
                    shpC5.Visible = True
                End If
                If fE Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCE10.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                        End If
                    dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKA10.Text)
                    txtTCA10(0).Text = CDbl(txtCSA10(0).Text) - CDbl(txtCRCA10(0).Text) + dPrezzoCrr + CDbl(txtEKA10.Text)
                    txtTSA10(0).Text = CDbl(txtCSSA10(0).Text) - CDbl(txtCRSA10(0).Text) + dPrezzoStd + CDbl(txtEKA10.Text)
                    txtCRSA10(0).Visible = True
                    txtCRCA10(0).Visible = True
                    txtCSA10(0).Visible = True
                    txtCSSA10(0).Visible = True
                    txtTCA10(0).Visible = True
                    txtTSA10(0).Visible = True
                End If



End Sub
Private Sub Sleeve()
'attivo il flag per le tagliatrici
                fS = True
                
                
                txtS01.Visible = True
                txtKHS01.Visible = True
                txtEHS01.Visible = True
                txtEKS01.Visible = True
                txtSS01.Visible = True
                txtRS01.Visible = True
                txtRCS01.Visible = True
                
                
                txtS01.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                    txtKHS01.Text = rsCiclo("Cicli/Ora")
                    txtEHS01.Text = rsCiclo("Costo€")
                    txtEKS01.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"
                
                Else
                    txtKHS01.Text = txtKHS01.Text
                    txtEHS01.Text = txtEHS01.Text
                    txtEKS01.Text = txtEHS01.Text / txtKHS01.Text
                    sScarto = "0"
                    sRecupero = "0"
               
                End If
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSS01.Text <> sScarto Then
                    If Not fRicalcola Then
                        txtSS01.Text = rsCiclo("Scarto")
                    Else
                        txtSS01.Text = txtSS01.Text
                    End If
                    If f1 Then
                        dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKS01.Text)
                        dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKS01.Text)
                        txtCSS01(0).Text = CalcolaScarto(txtSS01.Text, dPrezzoCrr)
                        txtCSSS01(0).Text = CalcolaScarto(txtSS01.Text, dPrezzoStd)
                    End If
                    If f2 Then
                        dPrezzoCrr = CDbl(txtTCC2.Text) ' + CDbl(txtEKS01.Text)
                        dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKS01.Text)
                        txtCSS01(1).Text = CalcolaScarto(txtSS01.Text, dPrezzoCrr)
                        txtCSSS01(1).Text = CalcolaScarto(txtSS01.Text, dPrezzoStd)
                    End If
                    If f3 Then
                        dPrezzoCrr = CDbl(txtTCC3.Text) ' + CDbl(txtEKS01.Text)
                        dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKS01.Text)
                        txtCSS01(2).Text = CalcolaScarto(txtSS01.Text, dPrezzoCrr)
                        txtCSSS01(2).Text = CalcolaScarto(txtSS01.Text, dPrezzoStd)
                    End If
                    If f5 Then
                        dPrezzoCrr = CDbl(txtTCC5.Text) ' + CDbl(txtEKS01.Text)
                        dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKS01.Text)
                        txtCSS01(3).Text = CalcolaScarto(txtSS01.Text, dPrezzoCrr)
                        txtCSSS01(3).Text = CalcolaScarto(txtSS01.Text, dPrezzoStd)
                    End If
                Else
                    
                    txtSS01.Text = "0"
                    For i = 0 To txtCSSS01.Count - 1
                        txtCSSS01(i).Text = "0"
                        txtCSS01(i).Text = "0"
                    Next i
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRS01.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRS01.Text = rsCiclo("Recupero")
                        txtRCS01.Text = rsCiclo("CRecupero")
                    Else
                        
                        txtRS01.Text = txtRS01.Text
                        txtRCS01.Text = txtRCS01.Text
                    End If
                    
                For i = 0 To txtCRSS01.Count - 1
                    txtCRSS01(i).Text = CDbl((txtSS01.Text) / 100) * CDbl((txtRS01.Text / 100)) * CDbl(txtRCS01.Text)
                    txtCRCS01(i).Text = CDbl((txtSS01.Text) / 100) * CDbl((txtRS01.Text / 100)) * CDbl(txtRCS01.Text)
                Next i
                Else
                    txtRS01.Text = "0"
                    For i = 0 To txtCRSS01.Count - 1
                        txtCRSS01(i).Text = "0"
                        txtCRCS01(i).Text = "0"
                    Next i
                End If
'   Totali
                
                
                If f1 Then
                    dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKS01.Text)
                    dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKS01.Text)
                    txtTCS01(0).Text = CDbl(txtCSS01(0).Text) - CDbl(txtCRCS01(0).Text) + dPrezzoCrr + CDbl(txtEKS01.Text)
                    txtTSS01(0).Text = CDbl(txtCSSS01(0).Text) - CDbl(txtCRSS01(0).Text) + dPrezzoStd + CDbl(txtEKS01.Text)
                    txtCRSS01(0).Visible = True
                    txtCRCS01(0).Visible = True
                    txtCSS01(0).Visible = True
                    txtCSSS01(0).Visible = True
                    txtTCS01(0).Visible = True
                    txtTSS01(0).Visible = True
                End If
                If f2 Then
                    dPrezzoCrr = CDbl(txtTCC2.Text) ' + CDbl(txtEKS01.Text)
                    dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKS01.Text)
                    txtTCS01(1).Text = CDbl(txtCSS01(1).Text) - CDbl(txtCRCS01(1).Text) + dPrezzoCrr + CDbl(txtEKS01.Text)
                    txtTSS01(1).Text = CDbl(txtCSSS01(1).Text) - CDbl(txtCRSS01(1).Text) + dPrezzoStd + CDbl(txtEKS01.Text)
                    Posizione txtCRSS01(1), True, txtCRSC2.Left, txtCRSS01(0).Top
                    Posizione txtCRCS01(1), True, txtCRCC2.Left, txtCRCS01(0).Top
                    Posizione txtCSS01(1), True, txtCSC2.Left, txtCSS01(0).Top
                    Posizione txtCSSS01(1), True, txtCSSC2.Left, txtCSSS01(0).Top
                    Posizione txtTCS01(1), True, txtTCC2.Left, txtTCS01(0).Top
                    Posizione txtTSS01(1), True, txtTSC2.Left, txtTSS01(0).Top
                    shpC2.Visible = True
                End If
                If f3 Then
                    dPrezzoCrr = CDbl(txtTCC3.Text) ' + CDbl(txtEKS01.Text)
                    dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKS01.Text)
                    txtTCS01(2).Text = CDbl(txtCSS01(2).Text) - CDbl(txtCRCS01(2).Text) + dPrezzoCrr + CDbl(txtEKS01.Text)
                    txtTSS01(2).Text = CDbl(txtCSSS01(2).Text) - CDbl(txtCRSS01(2).Text) + dPrezzoStd + CDbl(txtEKS01.Text)
                    Posizione txtCRSS01(2), True, txtCRSC3.Left, txtCRSS01(0).Top
                    Posizione txtCRCS01(2), True, txtCRCC3.Left, txtCRCS01(0).Top
                    Posizione txtCSS01(2), True, txtCSC3.Left, txtCSS01(0).Top
                    Posizione txtCSSS01(2), True, txtCSSC3.Left, txtCSSS01(0).Top
                    Posizione txtTCS01(2), True, txtTCC3.Left, txtTCS01(0).Top
                    Posizione txtTSS01(2), True, txtTSC3.Left, txtTSS01(0).Top
                    shpC3.Visible = True
                End If
                If f5 Then
                    
                    dPrezzoCrr = CDbl(txtTCC5.Text) ' + CDbl(txtEKS01.Text)
                    dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKS01.Text)
                    txtTCS01(3).Text = CDbl(txtCSS01(3).Text) - CDbl(txtCRCS01(3).Text) + dPrezzoCrr + CDbl(txtEKS01.Text)
                    txtTSS01(3).Text = CDbl(txtCSSS01(3).Text) - CDbl(txtCRSS01(3).Text) + dPrezzoStd + CDbl(txtEKS01.Text)
                    Posizione txtCRSS01(3), True, txtCRSC5.Left, txtCRSS01(0).Top
                    Posizione txtCRCS01(3), True, txtCRCC5.Left, txtCRCS01(0).Top
                    Posizione txtCSS01(3), True, txtCSC5.Left, txtCSS01(0).Top
                    Posizione txtCSSS01(3), True, txtCSSC5.Left, txtCSSS01(0).Top
                    Posizione txtTCS01(3), True, txtTCC5.Left, txtTCS01(0).Top
                    Posizione txtTSS01(3), True, txtTSC5.Left, txtTSS01(0).Top
                    shpC5.Visible = True
                End If
End Sub
Private Sub Squadratrice()
             
                txtQ01.Visible = True
                txtKHQ01.Visible = True
                txtEHQ01.Visible = True
                txtEKQ01.Visible = True
                txtSQ01.Visible = True
'                txtCSQ01.Visible = True
'                txtCSSQ01.Visible = True
                txtRQ01.Visible = True
                txtRCQ01.Visible = True
'                txtCRSQ01.Visible = True
'                txtCRCQ01.Visible = True
                
                txtQ01.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                    txtKHQ01.Text = rsCiclo("Cicli/Ora")
                    txtEHQ01.Text = rsCiclo("Costo€")
                    txtEKQ01.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"
               
                Else
                    txtKHQ01.Text = txtKHQ01.Text
                    txtEHQ01.Text = txtEHQ01.Text
                    txtEKQ01.Text = Format(CDbl(txtEHQ01.Text) / CDbl(txtKHQ01.Text), "0.00")
                    sScarto = "0"
                    sRecupero = "0"
               
                End If
                    If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSQ01.Text <> sScarto Then
                        If Not fRicalcola Then
                            txtSQ01.Text = rsCiclo("Scarto")
                        Else
                            txtSQ01.Text = txtSQ01.Text
                        End If
                        
                        If fS Or fA Or fM Then
                            If fS Then
                                For i = 0 To txtTCS01.Count - 1
                                    dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKbj.Text)
                                    dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKbj.Text)
                                    txtCSQ01(i).Text = CalcolaScarto(txtSQ01.Text, dPrezzoCrr)
                                    txtCSSQ01(i).Text = CalcolaScarto(txtSQ01.Text, dPrezzoStd)
                                Next i
                                
                            End If
                            If fM Then
                                For i = 0 To txtTCM01.Count - 1
                                    dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKQ01.Text)
                                    dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKQ01.Text)
                                    txtCSQ01(i).Text = CalcolaScarto(txtSQ01.Text, dPrezzoCrr)
                                    txtCSSQ01(i).Text = CalcolaScarto(txtSQ01.Text, dPrezzoStd)
                                Next i
                                
                            End If
                            If fA Then
                                For i = 0 To txtTCA10.Count - 1
                                    dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKQ01.Text)
                                    dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKQ01.Text)
                                    txtCSQ01(i).Text = CalcolaScarto(txtSQ01.Text, dPrezzoCrr)
                                    txtCSSQ01(i).Text = CalcolaScarto(txtSQ01.Text, dPrezzoStd)
                                Next i
                                
                            End If
                        Else
                            If f1 Then
                                dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKA10.Text)
                                dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKA10.Text)
                                txtCSQ01(0).Text = CalcolaScarto(txtSQ01.Text, dPrezzoCrr)
                                txtCSSQ01(0).Text = CalcolaScarto(txtSQ01.Text, dPrezzoStd)
                            End If
                            If f2 Then
                                dPrezzoCrr = CDbl(txtTCC2.Text) ' + CDbl(txtEKQ01.Text)
                                dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKQ01.Text)
                                txtCSQ01(1).Text = CalcolaScarto(txtSQ01.Text, dPrezzoCrr)
                                txtCSSQ01(1).Text = CalcolaScarto(txtSQ01.Text, dPrezzoStd)
                            End If
                            If f3 Then
                                dPrezzoCrr = CDbl(txtTCC3.Text) ' + CDbl(txtEKQ01.Text)
                                dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKQ01.Text)
                                txtCSQ01(2).Text = CalcolaScarto(txtSQ01.Text, dPrezzoCrr)
                                txtCSSQ01(2).Text = CalcolaScarto(txtSQ01.Text, dPrezzoStd)
                            End If
                            If f5 Then
                                dPrezzoCrr = CDbl(txtTCC5.Text) ' + CDbl(txtEKQ01.Text)
                                dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKQ01.Text)
                                txtCSQ01(3).Text = CalcolaScarto(txtSQ01.Text, dPrezzoCrr)
                                txtCSSQ01(3).Text = CalcolaScarto(txtSQ01.Text, dPrezzoStd)
                            End If
                        End If
                Else
                    txtSQ01.Text = "0"
                    For i = 0 To txtCSSQ01.Count - 1
                        txtCSSQ01(i).Text = "0"
                        txtCSQ01(i).Text = "0"
                    Next i
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRQ01.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRQ01.Text = rsCiclo("Recupero")
                      txtRCQ01.Text = rsCiclo("CRecupero")
                    Else
                        txtRQ01.Text = txtRQ01.Text
                        txtRCQ01.Text = txtRCQ01.Text
                    End If
                For i = 0 To txtCRSQ01.Count - 1
                    txtCRSQ01(i).Text = CDbl((txtSQ01.Text) / 100) * CDbl((txtRQ01.Text / 100)) * CDbl(txtRCQ01.Text)
                    txtCRCQ01(i).Text = CDbl((txtSQ01.Text) / 100) * CDbl((txtRQ01.Text / 100)) * CDbl(txtRCQ01.Text)
                Next i
                Else
                    txtRQ01.Text = "0"
                    For i = 0 To txtCRSQ01.Count - 1
                        txtCRSQ01(i).Text = "0"
                        txtCRCQ01(i).Text = "0"
                    Next i
                End If
'   Totali
                
                If fS Or fA Or fM Then
                    If fS Then
                        For i = 0 To txtTCS01.Count - 1
                            dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKQ01.Text)
                            dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKQ01.Text)
                            txtTCQ01(i).Text = CDbl(txtCSQ01(i).Text) - CDbl(txtCRCQ01(i).Text) + dPrezzoCrr + CDbl(txtEKQ01.Text)
                            txtTSQ01(i).Text = CDbl(txtCSSQ01(i).Text) - CDbl(txtCRSQ01(i).Text) + dPrezzoStd + CDbl(txtEKQ01.Text)
                            Posizione txtCRSQ01(i), True, txtCRSS01(i).Left, txtCRSQ01(0).Top
                            Posizione txtCRCQ01(i), True, txtCRCS01(i).Left, txtCRCQ01(0).Top
                            Posizione txtCSQ01(i), True, txtCSS01(i).Left, txtCSQ01(0).Top
                            Posizione txtCSSQ01(i), True, txtCSSS01(i).Left, txtCSSQ01(0).Top
                            Posizione txtTCQ01(i), True, txtTCS01(i).Left, txtTCQ01(0).Top
                            Posizione txtTSQ01(i), True, txtTSS01(i).Left, txtTSQ01(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fM Then
                        For i = 0 To txtTCM01.Count - 1
                            dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKQ01.Text)
                            dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKQ01.Text)
                            txtTCQ01(i).Text = CDbl(txtCSQ01(i).Text) - CDbl(txtCRCQ01(i).Text) + dPrezzoCrr + CDbl(txtEKQ01.Text)
                            txtTSQ01(i).Text = CDbl(txtCSSQ01(i).Text) - CDbl(txtCRSQ01(i).Text) + dPrezzoStd + CDbl(txtEKQ01.Text)
                            Posizione txtCRSQ01(i), True, txtCRSM01(i).Left, txtCRSQ01(0).Top
                            Posizione txtCRCQ01(i), True, txtCRCM01(i).Left, txtCRCQ01(0).Top
                            Posizione txtCSQ01(i), True, txtCSM01(i).Left, txtCSQ01(0).Top
                            Posizione txtCSSQ01(i), True, txtCSSM01(i).Left, txtCSSQ01(0).Top
                            Posizione txtTCQ01(i), True, txtTCM01(i).Left, txtTCQ01(0).Top
                            Posizione txtTSQ01(i), True, txtTSM01(i).Left, txtTSQ01(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fA Then
                        For i = 0 To txtTCA10.Count - 1
                            dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKQ01.Text)
                            dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKQ01.Text)
                            txtTCQ01(i).Text = CDbl(txtCSQ01(i).Text) - CDbl(txtCRCQ01(i).Text) + dPrezzoCrr + CDbl(txtEKQ01.Text)
                            txtTSQ01(i).Text = CDbl(txtCSSQ01(i).Text) - CDbl(txtCRSQ01(i).Text) + dPrezzoStd + CDbl(txtEKQ01.Text)
                            Posizione txtCRSQ01(i), True, txtCRSA10(i).Left, txtCRSQ01(0).Top
                            Posizione txtCRCQ01(i), True, txtCRCA10(i).Left, txtCRCQ01(0).Top
                            Posizione txtCSQ01(i), True, txtCSA10(i).Left, txtCSQ01(0).Top
                            Posizione txtCSSQ01(i), True, txtCSSA10(i).Left, txtCSSQ01(0).Top
                            Posizione txtTCQ01(i), True, txtTCA10(i).Left, txtTCQ01(0).Top
                            Posizione txtTSQ01(i), True, txtTSA10(i).Left, txtTSQ01(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
' totali senza lavorazioni intermedie
                Else
                    If f1 Then
                        dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKS01.Text)
                        dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKS01.Text)
                        txtTCQ01(0).Text = CDbl(txtCSQ01(0).Text) - CDbl(txtCRCQ01(0).Text) + dPrezzoCrr + CDbl(txtEKQ01.Text)
                        txtTSQ01(0).Text = CDbl(txtCSSQ01(0).Text) - CDbl(txtCRSQ01(0).Text) + dPrezzoStd + CDbl(txtEKQ01.Text)
                        txtCRSQ01(0).Visible = True
                        txtCRCQ01(0).Visible = True
                        txtCSQ01(0).Visible = True
                        txtCSSQ01(0).Visible = True
                        txtTCQ01(0).Visible = True
                        txtTSQ01(0).Visible = True
                    End If
                    If f2 Then
                        dPrezzoCrr = CDbl(txtTCC2.Text) ' + CDbl(txtEKQ01.Text)
                        dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKQ01.Text)
                        txtTCQ01(1).Text = CDbl(txtCSQ01(1).Text) - CDbl(txtCRCQ01(1).Text) + dPrezzoCrr + CDbl(txtEKQ01.Text)
                        txtTSQ01(1).Text = CDbl(txtCSSQ01(1).Text) - CDbl(txtCRSQ01(1).Text) + dPrezzoStd + CDbl(txtEKQ01.Text)
                        Posizione txtCRSQ01(1), True, txtCRSC2.Left, txtCRSQ01(0).Top
                        Posizione txtCRCQ01(1), True, txtCRCC2.Left, txtCRCQ01(0).Top
                        Posizione txtCSQ01(1), True, txtCSC2.Left, txtCSQ01(0).Top
                        Posizione txtCSSQ01(1), True, txtCSSC2.Left, txtCSSQ01(0).Top
                        Posizione txtTCQ01(1), True, txtTCC2.Left, txtTCQ01(0).Top
                        Posizione txtTSQ01(1), True, txtTSC2.Left, txtTSQ01(0).Top
                        shpC2.Visible = True
                    End If
                    If f3 Then
                        dPrezzoCrr = CDbl(txtTCC3.Text) ' + CDbl(txtEKQ01.Text)
                        dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKQ01.Text)
                        txtTCQ01(2).Text = CDbl(txtCSQ01(2).Text) - CDbl(txtCRCQ01(2).Text) + dPrezzoCrr + CDbl(txtEKQ01.Text)
                        txtTSQ01(2).Text = CDbl(txtCSSQ01(2).Text) - CDbl(txtCRSQ01(2).Text) + dPrezzoStd + CDbl(txtEKQ01.Text)
                        Posizione txtCRSQ01(2), True, txtCRSC3.Left, txtCRSQ01(0).Top
                        Posizione txtCRCQ01(2), True, txtCRCC3.Left, txtCRCQ01(0).Top
                        Posizione txtCSQ01(2), True, txtCSC3.Left, txtCSQ01(0).Top
                        Posizione txtCSSQ01(2), True, txtCSSC3.Left, txtCSSQ01(0).Top
                        Posizione txtTCQ01(2), True, txtTCC3.Left, txtTCQ01(0).Top
                        Posizione txtTSQ01(2), True, txtTSC3.Left, txtTSQ01(0).Top
                        shpC3.Visible = True
                    End If
                    If f5 Then
                        dPrezzoCrr = CDbl(txtTCC5.Text) ' + CDbl(txtEKQ01.Text)
                        dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKQ01.Text)
                        txtTCQ01(3).Text = CDbl(txtCSQ01(3).Text) - CDbl(txtCRCQ01(3).Text) + dPrezzoCrr + CDbl(txtEKQ01.Text)
                        txtTSQ01(3).Text = CDbl(txtCSSQ01(3).Text) - CDbl(txtCRSQ01(3).Text) + dPrezzoStd + CDbl(txtEKQ01.Text)
                        Posizione txtCRSQ01(3), True, txtCRSC5.Left, txtCRSQ01(0).Top
                        Posizione txtCRCQ01(3), True, txtCRCC5.Left, txtCRCQ01(0).Top
                        Posizione txtCSQ01(3), True, txtCSC5.Left, txtCSQ01(0).Top
                        Posizione txtCSSQ01(3), True, txtCSSC5.Left, txtCSSQ01(0).Top
                        Posizione txtTCQ01(3), True, txtTCC5.Left, txtTCQ01(0).Top
                        Posizione txtTSQ01(3), True, txtTSC5.Left, txtTSQ01(0).Top
                        shpC5.Visible = True
                    End If
                End If
End Sub


Private Sub B00()

                
                txtB00.Visible = True
                txtKHB00.Visible = True
                txtEHB00.Visible = True
                txtEKB00.Visible = True
                txtSB00.Visible = True
'                txtCSB00.Visible = True
'                txtCSSB00.Visible = True
                txtRB00.Visible = True
                txtRCB00.Visible = True
'                txtCRSB00.Visible = True
'                txtCRCB00.Visible = True
                
                txtB00.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                    txtKHB00.Text = rsCiclo("Cicli/Ora")
                    txtEHB00.Text = rsCiclo("Costo€")
                    txtEKB00.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"
                    
                Else
                    txtKHB00.Text = txtKHB00.Text
                    txtEHB00.Text = txtEHB00.Text
                    txtEKB00.Text = txtEHB00.Text / txtKHB00.Text
                    sScarto = "0"
                    sRecupero = "0"
                End If
                If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSB00.Text <> sScarto Then
                    If Not fRicalcola Then
                        txtSB00.Text = rsCiclo("Scarto")
                    Else
                        txtSB00.Text = txtSB00.Text
                    End If
                
                    If fS Or fA Or fM Then
                        If fS Then
                            For i = 0 To txtTCS01.Count - 1
                                dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKB00.Text)
                                dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKB00.Text)
                                txtCSB00(i).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                                txtCSSB00(i).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                            Next i
                        End If
                        If fM Then
                            If Not fE Then
                                iLinea = txtTCM01.Count - 1
                            End If
                
                            For i = 0 To iLinea
                                dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKB00.Text)
                                dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKB00.Text)
                                txtCSB00(i).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                                txtCSSB00(i).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                            Next i
                        End If
                        If fA Then
                            If Not fE Then
                                iLinea = txtTCA10.Count - 1
                            End If
                
                            For i = 0 To iLinea
                                dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKB00.Text)
                                dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKB00.Text)
                                txtCSB00(i).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                                txtCSSB00(i).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                            Next i
                        End If
                    Else
                        If f1 Then
                            If Not fAcc Then    ' Controllo se accoppiato
                                dPrezzoCrr = CDbl(txtTCC1.Text)
                            Else
                                dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                            End If
                            dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKA10.Text)
                            txtCSB00(0).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                            txtCSSB00(0).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                        End If
                        If f2 Then
                            If Not fAcc Then    ' Controllo se accoppiato
                                dPrezzoCrr = CDbl(txtTCC2.Text)
                            Else
                                dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                            End If
                            dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKb00.Text)
                            txtCSB00(1).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                            txtCSSB00(1).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                        End If
                        If f3 Then
                            If Not fAcc Then    ' Controllo se accoppiato
                                dPrezzoCrr = CDbl(txtTCC3.Text)
                            Else
                                dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                            End If
                            dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKb00.Text)
                            txtCSB00(2).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                            txtCSSB00(2).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                        End If
                        If f5 Then
                            If Not fAcc Then    ' Controllo se accoppiato
                                dPrezzoCrr = CDbl(txtTCC5.Text)
                            Else
                                dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                            End If
                            dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKb00.Text)
                            txtCSB00(3).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                            txtCSSB00(3).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                        End If
                        If fE Then
                            If Not fAcc Then    ' Controllo se accoppiato
                                dPrezzoCrr = CDbl(txtTCE10.Text)
                            Else
                                dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                            End If
                            dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKA10.Text)
                            txtCSB00(0).Text = CalcolaScarto(txtSB00.Text, dPrezzoCrr)
                            txtCSSB00(0).Text = CalcolaScarto(txtSB00.Text, dPrezzoStd)
                        End If
                
                    End If
                Else
                    txtSB00.Text = "0"
                    For i = 0 To txtCSSB00.Count - 1
                        txtCSSB00(i).Text = "0"
                        txtCSB00(i).Text = "0"
                    Next i
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRB00.Text <> sRecupero Then
                    If Not fRicalcola Then
                        txtRB00.Text = rsCiclo("Recupero")
                        txtRCB00.Text = rsCiclo("CRecupero")
                    Else
                        txtRB00.Text = txtRB00.Text
                        txtRCB00.Text = txtRCB00.Text
                    End If
                    For i = 0 To txtCRSB00.Count - 1
                        txtCRSB00(i).Text = CDbl((txtSB00.Text) / 100) * CDbl((txtRB00.Text / 100)) * CDbl(txtRCB00.Text)
                        txtCRCB00(i).Text = CDbl((txtSB00.Text) / 100) * CDbl((txtRB00.Text / 100)) * CDbl(txtRCB00.Text)
                    Next i
                Else
                    txtRB00.Text = "0"
                    For i = 0 To txtCRSB00.Count - 1
                        txtCRSB00(i).Text = "0"
                        txtCRCB00(i).Text = "0"
                    Next i
                End If
'   Totali
                
                If fS Or fA Or fM Then
                    If fS Then
                        For i = 0 To txtTCS01.Count - 1
                            dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKB00.Text)
                            dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKB00.Text)
                            txtTCB00(i).Text = CDbl(txtCSB00(i).Text) - CDbl(txtCRCB00(i).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                            txtTSB00(i).Text = CDbl(txtCSSB00(i).Text) - CDbl(txtCRSB00(i).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                            Posizione txtCRSB00(i), True, txtCRSS01(i).Left, txtCRSB00(0).Top
                            Posizione txtCRCB00(i), True, txtCRCS01(i).Left, txtCRCB00(0).Top
                            Posizione txtCSB00(i), True, txtCSS01(i).Left, txtCSB00(0).Top
                            Posizione txtCSSB00(i), True, txtCSSS01(i).Left, txtCSSB00(0).Top
                            Posizione txtTCB00(i), True, txtTCS01(i).Left, txtTCB00(0).Top
                            Posizione txtTSB00(i), True, txtTSS01(i).Left, txtTSB00(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fM Then
                        If Not fE Then
                            iLinea = txtTCM01.Count - 1
                        End If
                        For i = 0 To iLinea 'txtTCM01.Count - 1
                            dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKB00.Text)
                            dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKB00.Text)
                            txtTCB00(i).Text = CDbl(txtCSB00(i).Text) - CDbl(txtCRCB00(i).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                            txtTSB00(i).Text = CDbl(txtCSSB00(i).Text) - CDbl(txtCRSB00(i).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                            Posizione txtCRSB00(i), True, txtCRSM01(i).Left, txtCRSB00(0).Top
                            Posizione txtCRCB00(i), True, txtCRCM01(i).Left, txtCRCB00(0).Top
                            Posizione txtCSB00(i), True, txtCSM01(i).Left, txtCSB00(0).Top
                            Posizione txtCSSB00(i), True, txtCSSM01(i).Left, txtCSSB00(0).Top
                            Posizione txtTCB00(i), True, txtTCM01(i).Left, txtTCB00(0).Top
                            Posizione txtTSB00(i), True, txtTSM01(i).Left, txtTSB00(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fA Then
                        If Not fE Then
                            iLinea = txtTCA10.Count - 1
                        End If
            
                        For i = 0 To iLinea
                            dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKB00.Text)
                            dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKB00.Text)
                            txtTCB00(i).Text = CDbl(txtCSB00(i).Text) - CDbl(txtCRCB00(i).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                            txtTSB00(i).Text = CDbl(txtCSSB00(i).Text) - CDbl(txtCRSB00(i).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                            Posizione txtCRSB00(i), True, txtCRSA10(i).Left, txtCRSB00(0).Top
                            Posizione txtCRCB00(i), True, txtCRCA10(i).Left, txtCRCB00(0).Top
                            Posizione txtCSB00(i), True, txtCSA10(i).Left, txtCSB00(0).Top
                            Posizione txtCSSB00(i), True, txtCSSA10(i).Left, txtCSSB00(0).Top
                            Posizione txtTCB00(i), True, txtTCA10(i).Left, txtTCB00(0).Top
                            Posizione txtTSB00(i), True, txtTSA10(i).Left, txtTSB00(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If

' totali senza lavorazioni intermedie
                Else
                    If f1 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC1.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKS01.Text)
                        txtTCB00(0).Text = CDbl(txtCSB00(0).Text) - CDbl(txtCRCB00(0).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                        txtTSB00(0).Text = CDbl(txtCSSB00(0).Text) - CDbl(txtCRSB00(0).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                        txtCRSB00(0).Visible = True
                        txtCRCB00(0).Visible = True
                        txtCSB00(0).Visible = True
                        txtCSSB00(0).Visible = True
                        txtTCB00(0).Visible = True
                        txtTSB00(0).Visible = True
                    End If
                    If f2 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC2.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKb00.Text)
                        txtTCB00(1).Text = CDbl(txtCSB00(1).Text) - CDbl(txtCRCB00(1).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                        txtTSB00(1).Text = CDbl(txtCSSB00(1).Text) - CDbl(txtCRSB00(1).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                        Posizione txtCRSB00(1), True, txtCRSC2.Left, txtCRSB00(0).Top
                        Posizione txtCRCB00(1), True, txtCRCC2.Left, txtCRCB00(0).Top
                        Posizione txtCSB00(1), True, txtCSC2.Left, txtCSB00(0).Top
                        Posizione txtCSSB00(1), True, txtCSSC2.Left, txtCSSB00(0).Top
                        Posizione txtTCB00(1), True, txtTCC2.Left, txtTCB00(0).Top
                        Posizione txtTSB00(1), True, txtTSC2.Left, txtTSB00(0).Top
                        shpC2.Visible = True
                    End If
                    If f3 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC3.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKb00.Text)
                        txtTCB00(2).Text = CDbl(txtCSB00(2).Text) - CDbl(txtCRCB00(2).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                        txtTSB00(2).Text = CDbl(txtCSSB00(2).Text) - CDbl(txtCRSB00(2).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                        Posizione txtCRSB00(2), True, txtCRSC3.Left, txtCRSB00(0).Top
                        Posizione txtCRCB00(2), True, txtCRCC3.Left, txtCRCB00(0).Top
                        Posizione txtCSB00(2), True, txtCSC3.Left, txtCSB00(0).Top
                        Posizione txtCSSB00(2), True, txtCSSC3.Left, txtCSSB00(0).Top
                        Posizione txtTCB00(2), True, txtTCC3.Left, txtTCB00(0).Top
                        Posizione txtTSB00(2), True, txtTSC3.Left, txtTSB00(0).Top
                        shpC3.Visible = True
                    End If
                    If f5 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC5.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKb00.Text)
                        txtTCB00(3).Text = CDbl(txtCSB00(3).Text) - CDbl(txtCRCB00(3).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                        txtTSB00(3).Text = CDbl(txtCSSB00(3).Text) - CDbl(txtCRSB00(3).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                        Posizione txtCRSB00(3), True, txtCRSC5.Left, txtCRSB00(0).Top
                        Posizione txtCRCB00(3), True, txtCRCC5.Left, txtCRCB00(0).Top
                        Posizione txtCSB00(3), True, txtCSC5.Left, txtCSB00(0).Top
                        Posizione txtCSSB00(3), True, txtCSSC5.Left, txtCSSB00(0).Top
                        Posizione txtTCB00(3), True, txtTCC5.Left, txtTCB00(0).Top
                        Posizione txtTSB00(3), True, txtTSC5.Left, txtTSB00(0).Top
                        shpC5.Visible = True
                    End If
                    If fE Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCE10.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKS01.Text)
                        txtTCB00(0).Text = CDbl(txtCSB00(0).Text) - CDbl(txtCRCB00(0).Text) + dPrezzoCrr + CDbl(txtEKB00.Text)
                        txtTSB00(0).Text = CDbl(txtCSSB00(0).Text) - CDbl(txtCRSB00(0).Text) + dPrezzoStd + CDbl(txtEKB00.Text)
                        txtCRSB00(0).Visible = True
                        txtCRCB00(0).Visible = True
                        txtCSB00(0).Visible = True
                        txtCSSB00(0).Visible = True
                        txtTCB00(0).Visible = True
                        txtTSB00(0).Visible = True
                    End If

                
                End If
            


End Sub

Private Sub BJ()
                
                
                txtBJ.Visible = True
                txtKHBJ.Visible = True
                txtEHBJ.Visible = True
                txtEKBJ.Visible = True
                txtSBJ.Visible = True
'                txtCSBJ.Visible = True
'                txtCSSBJ.Visible = True
                txtRBJ.Visible = True
                txtRCBJ.Visible = True
'                txtCRSBJ.Visible = True
'                txtCRCBJ.Visible = True
                
                txtBJ.Text = rsCiclo("Descrizione")
                        If Not fRicalcola Then
                            txtKHBJ.Text = rsCiclo("Cicli/Ora")
                            txtEHBJ.Text = rsCiclo("Costo€")
                            txtEKBJ.Text = rsCiclo("CostoKG€")
                            sScarto = "e/kg"
                            sRecupero = "e/kg"
                        Else
                            txtKHBJ.Text = txtKHBJ.Text
                            txtEHBJ.Text = txtEHBJ.Text
                            txtEKBJ.Text = txtEHBJ.Text / txtKHBJ.Text
                            txtSBJ.Text = txtSBJ.Text
                            sScarto = "0"
                            sRecupero = "0"
                        End If
                
       '                 If Not fRicalcola Then
'                    If Not txtSBJ.Text = "0" Then
                        If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSBJ.Text <> sScarto Then
                            If Not fRicalcola Then
                                txtSBJ.Text = rsCiclo("Scarto")
                            Else
                                txtSBJ.Text = txtSBJ.Text
                            End If
                            If fS Or fA Or fM Then
                                If fS Then
                                    For i = 0 To txtTCS01.Count - 1
                                        dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKbj.Text)
                                        dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKbj.Text)
                                        txtCSBJ(i).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                        txtCSSBJ(i).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                    Next i
                                    
                                End If
                                If fM Then
                                    If Not fE Then
                                        iLinea = txtTCM01.Count - 1
                                    End If
                                    For i = 0 To iLinea
                                        dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKbj.Text)
                                        dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKbj.Text)
                                        txtCSBJ(i).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                        txtCSSBJ(i).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                    Next i
                                    
                                End If
                                If fA Then
                                    If Not fE Then
                                        iLinea = txtTCA10.Count - 1
                                    End If
                        
                                    For i = 0 To iLinea
                                        dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKbj.Text)
                                        dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKbj.Text)
                                        txtCSBJ(i).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                        txtCSSBJ(i).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                    Next i
                                    
                                End If
                            Else
                                If f1 Then
                                    If Not fAcc Then    ' Controllo se accoppiato
                                        dPrezzoCrr = CDbl(txtTCC1.Text)
                                    Else
                                        dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                                    End If
                                    dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKA10.Text)
                                    txtCSBJ(0).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                    txtCSSBJ(0).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                End If
                                If f2 Then
                                    If Not fAcc Then    ' Controllo se accoppiato
                                        dPrezzoCrr = CDbl(txtTCC2.Text)
                                    Else
                                        dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                                    End If
                                    dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKbj.Text)
                                    txtCSBJ(1).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                    txtCSSBJ(1).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                End If
                                If f3 Then
                                    If Not fAcc Then    ' Controllo se accoppiato
                                        dPrezzoCrr = CDbl(txtTCC3.Text)
                                    Else
                                        dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                                    End If
                                    dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKbj.Text)
                                    txtCSBJ(2).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                    txtCSSBJ(2).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                End If
                                If f5 Then
                                    If Not fAcc Then    ' Controllo se accoppiato
                                        dPrezzoCrr = CDbl(txtTCC5.Text)
                                    Else
                                        dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                                    End If
                                    dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKbj.Text)
                                    txtCSBJ(3).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                    txtCSSBJ(3).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                End If
                                If fE Then
                                    If Not fAcc Then    ' Controllo se accoppiato
                                        dPrezzoCrr = CDbl(txtTCE10.Text)
                                    Else
                                        dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                                    End If
                                    dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKbj.Text)
                                    txtCSBJ(0).Text = CalcolaScarto(txtSBJ.Text, dPrezzoCrr)
                                    txtCSSBJ(0).Text = CalcolaScarto(txtSBJ.Text, dPrezzoStd)
                                End If
                            End If
                    Else
                        txtSBJ.Text = "0"
                        For i = 0 To txtCSSBJ.Count - 1
                            txtCSSBJ(i).Text = "0"
                            txtCSBJ(i).Text = "0"
                        Next i
                    End If
'                End If
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRBJ.Text <> sRecupero Then
                        If Not fRicalcola Then
                            txtRBJ.Text = rsCiclo("Recupero")
                            txtRCBJ.Text = rsCiclo("CRecupero")

                        Else
                            txtRBJ.Text = txtRBJ.Text
                            txtRCBJ.Text = txtRCBJ.Text
                        End If

                For i = 0 To txtCRSBJ.Count - 1
                    txtCRSBJ(i).Text = CDbl((txtSBJ.Text) / 100) * CDbl((txtRBJ.Text / 100)) * CDbl(txtRCBJ.Text)
                    txtCRCBJ(i).Text = CDbl((txtSBJ.Text) / 100) * CDbl((txtRBJ.Text / 100)) * CDbl(txtRCBJ.Text)
                Next i
                Else
                    txtRBJ.Text = "0"
                    For i = 0 To txtCRSBJ.Count - 1
                        txtCRSBJ(i).Text = "0"
                        txtCRCBJ(i).Text = "0"
                    Next i
                End If
'   Totali
                
                If fS Or fA Or fM Then
                    If fS Then
                        For i = 0 To txtTCS01.Count - 1
                            dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKbj.Text)
                            dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKbj.Text)
                            txtTCBJ(i).Text = CDbl(txtCSBJ(i).Text) - CDbl(txtCRCBJ(i).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                            txtTSBJ(i).Text = CDbl(txtCSSBJ(i).Text) - CDbl(txtCRSBJ(i).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                            Posizione txtCRSBJ(i), True, txtCRSS01(i).Left, txtCRSBJ(0).Top
                            Posizione txtCRCBJ(i), True, txtCRCS01(i).Left, txtCRCBJ(0).Top
                            Posizione txtCSBJ(i), True, txtCSS01(i).Left, txtCSBJ(0).Top
                            Posizione txtCSSBJ(i), True, txtCSSS01(i).Left, txtCSSBJ(0).Top
                            Posizione txtTCBJ(i), True, txtTCS01(i).Left, txtTCBJ(0).Top
                            Posizione txtTSBJ(i), True, txtTSS01(i).Left, txtTSBJ(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fM Then
                        If Not fE Then
                            iLinea = txtTCM01.Count - 1
                        End If
                        For i = 0 To iLinea
                            dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKbj.Text)
                            dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKbj.Text)
                            txtTCBJ(i).Text = CDbl(txtCSBJ(i).Text) - CDbl(txtCRCBJ(i).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                            txtTSBJ(i).Text = CDbl(txtCSSBJ(i).Text) - CDbl(txtCRSBJ(i).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                            Posizione txtCRSBJ(i), True, txtCRSM01(i).Left, txtCRSBJ(0).Top
                            Posizione txtCRCBJ(i), True, txtCRCM01(i).Left, txtCRCBJ(0).Top
                            Posizione txtCSBJ(i), True, txtCSM01(i).Left, txtCSBJ(0).Top
                            Posizione txtCSSBJ(i), True, txtCSSM01(i).Left, txtCSSBJ(0).Top
                            Posizione txtTCBJ(i), True, txtTCM01(i).Left, txtTCBJ(0).Top
                            Posizione txtTSBJ(i), True, txtTSM01(i).Left, txtTSBJ(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fA Then
                        If Not fE Then
                            iLinea = txtTCA10.Count - 1
                        End If
            
                        For i = 0 To iLinea

                            dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKbj.Text)
                            dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKbj.Text)
                            txtTCBJ(i).Text = CDbl(txtCSBJ(i).Text) - CDbl(txtCRCBJ(i).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                            txtTSBJ(i).Text = CDbl(txtCSSBJ(i).Text) - CDbl(txtCRSBJ(i).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                            Posizione txtCRSBJ(i), True, txtCRSA10(i).Left, txtCRSBJ(0).Top
                            Posizione txtCRCBJ(i), True, txtCRCA10(i).Left, txtCRCBJ(0).Top
                            Posizione txtCSBJ(i), True, txtCSA10(i).Left, txtCSBJ(0).Top
                            Posizione txtCSSBJ(i), True, txtCSSA10(i).Left, txtCSSBJ(0).Top
                            Posizione txtTCBJ(i), True, txtTCA10(i).Left, txtTCBJ(0).Top
                            Posizione txtTSBJ(i), True, txtTSA10(i).Left, txtTSBJ(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
' totali senza lavorazioni intermedie
                Else
                    If f1 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC1.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC1.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKS01.Text)
                        txtTCBJ(0).Text = CDbl(txtCSBJ(0).Text) - CDbl(txtCRCBJ(0).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                        txtTSBJ(0).Text = CDbl(txtCSSBJ(0).Text) - CDbl(txtCRSBJ(0).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                        txtCRSBJ(0).Visible = True
                        txtCRCBJ(0).Visible = True
                        txtCSBJ(0).Visible = True
                        txtCSSBJ(0).Visible = True
                        txtTCBJ(0).Visible = True
                        txtTSBJ(0).Visible = True
                    End If
                    If f2 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC2.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC2.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKbj.Text)
                        txtTCBJ(1).Text = CDbl(txtCSBJ(1).Text) - CDbl(txtCRCBJ(1).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                        txtTSBJ(1).Text = CDbl(txtCSSBJ(1).Text) - CDbl(txtCRSBJ(1).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                        Posizione txtCRSBJ(1), True, txtCRSC2.Left, txtCRSBJ(0).Top
                        Posizione txtCRCBJ(1), True, txtCRCC2.Left, txtCRCBJ(0).Top
                        Posizione txtCSBJ(1), True, txtCSC2.Left, txtCSBJ(0).Top
                        Posizione txtCSSBJ(1), True, txtCSSC2.Left, txtCSSBJ(0).Top
                        Posizione txtTCBJ(1), True, txtTCC2.Left, txtTCBJ(0).Top
                        Posizione txtTSBJ(1), True, txtTSC2.Left, txtTSBJ(0).Top
                        shpC2.Visible = True
                    End If
                    If f3 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC3.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC3.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKbj.Text)
                        txtTCBJ(2).Text = CDbl(txtCSBJ(2).Text) - CDbl(txtCRCBJ(2).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                        txtTSBJ(2).Text = CDbl(txtCSSBJ(2).Text) - CDbl(txtCRSBJ(2).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                        Posizione txtCRSBJ(2), True, txtCRSC3.Left, txtCRSBJ(0).Top
                        Posizione txtCRCBJ(2), True, txtCRCC3.Left, txtCRCBJ(0).Top
                        Posizione txtCSBJ(2), True, txtCSC3.Left, txtCSBJ(0).Top
                        Posizione txtCSSBJ(2), True, txtCSSC3.Left, txtCSSBJ(0).Top
                        Posizione txtTCBJ(2), True, txtTCC3.Left, txtTCBJ(0).Top
                        Posizione txtTSBJ(2), True, txtTSC3.Left, txtTSBJ(0).Top
                        shpC3.Visible = True
                    End If
                    If f5 Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCC5.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCC5.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKbj.Text)
                        txtTCBJ(3).Text = CDbl(txtCSBJ(3).Text) - CDbl(txtCRCBJ(3).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                        txtTSBJ(3).Text = CDbl(txtCSSBJ(3).Text) - CDbl(txtCRSBJ(3).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                        Posizione txtCRSBJ(3), True, txtCRSC5.Left, txtCRSBJ(0).Top
                        Posizione txtCRCBJ(3), True, txtCRCC5.Left, txtCRCBJ(0).Top
                        Posizione txtCSBJ(3), True, txtCSC5.Left, txtCSBJ(0).Top
                        Posizione txtCSSBJ(3), True, txtCSSC5.Left, txtCSSBJ(0).Top
                        Posizione txtTCBJ(3), True, txtTCC5.Left, txtTCBJ(0).Top
                        Posizione txtTSBJ(3), True, txtTSC5.Left, txtTSBJ(0).Top
                        shpC5.Visible = True
                    End If
                    If fE Then
                        If Not fAcc Then    ' Controllo se accoppiato
                            dPrezzoCrr = CDbl(txtTCE10.Text)
                        Else
                            dPrezzoCrr = CDbl(txtTCE10.Text) + CalcolaAccoppiati
                        End If
                        dPrezzoStd = CDbl(txtTSE10.Text) ' + CDbl(txtEKS01.Text)
                        txtTCBJ(0).Text = CDbl(txtCSBJ(0).Text) - CDbl(txtCRCBJ(0).Text) + dPrezzoCrr + CDbl(txtEKBJ.Text)
                        txtTSBJ(0).Text = CDbl(txtCSSBJ(0).Text) - CDbl(txtCRSBJ(0).Text) + dPrezzoStd + CDbl(txtEKBJ.Text)
                        txtCRSBJ(0).Visible = True
                        txtCRCBJ(0).Visible = True
                        txtCSBJ(0).Visible = True
                        txtCSSBJ(0).Visible = True
                        txtTCBJ(0).Visible = True
                        txtTSBJ(0).Visible = True
                    End If

                
                End If


End Sub
Private Sub B11()
               
                txtB11.Top = txtQ01.Top
                txtKHB11.Top = txtKHQ01.Top
                txtEHB11.Top = txtEHQ01.Top
                txtEKB11.Top = txtEKQ01.Top
                txtSB11.Top = txtSQ01.Top
                txtRB11.Top = txtRQ01.Top
                txtRCB11.Top = txtRCQ01.Top
                txtCSB11(0).Top = txtCSQ01(0).Top
                txtCSSB11(0).Top = txtCSSQ01(0).Top
                txtCRSB11(0).Top = txtCRSQ01(0).Top
                txtCRCB11(0).Top = txtCRCQ01(0).Top
                txtTSB11(0).Top = txtTSQ01(0).Top
                txtTCB11(0).Top = txtTCQ01(0).Top
                
                
                
                txtB11.Left = txtQ01.Left
                txtKHB11.Left = txtKHQ01.Left
                txtEHB11.Left = txtEHQ01.Left
                txtEKB11.Left = txtEKQ01.Left
                txtSB11.Left = txtSQ01.Left
                txtRB11.Left = txtRQ01.Left
                txtRCB11.Left = txtRCQ01.Left
                txtCSB11(0).Left = txtCSQ01(0).Left
                txtCSSB11(0).Left = txtCSSQ01(0).Left
                txtCRSB11(0).Left = txtCRSQ01(0).Left
                txtCRCB11(0).Left = txtCRCQ01(0).Left
                txtTSB11(0).Left = txtTSQ01(0).Left
                txtTCB11(0).Left = txtTCQ01(0).Left
                
                
                txtB11.Visible = True
                txtKHB11.Visible = True
                txtEHB11.Visible = True
                txtEKB11.Visible = True
                txtSB11.Visible = True
'                txtCSb11.Visible = True
'                txtCSSb11.Visible = True
                txtRB11.Visible = True
                txtRCB11.Visible = True
'                txtCRSb11.Visible = True
'                txtCRCb11.Visible = True
                
                txtB11.Text = rsCiclo("Descrizione")
                If Not fRicalcola Then
                    txtKHB11.Text = rsCiclo("Cicli/Ora")
                    txtEHB11.Text = rsCiclo("Costo€")
                    txtEKB11.Text = rsCiclo("CostoKG€")
                    sScarto = "e/kg"
                    sRecupero = "e/kg"

                Else
                    txtKHB11.Text = txtKHB11.Text
                    txtEHB11.Text = txtEHB11.Text
                    txtEKB11.Text = txtEHB11.Text / txtKHB11.Text
                    sScarto = "0"
                    sRecupero = "0"
                End If
                
                    If Not IsNull(rsCiclo("Scarto")) And rsCiclo("Scarto") <> "" Or txtSB11.Text <> sScarto Then
                    txtSB11.Text = rsCiclo("Scarto")
                        If fS Or fA Or fM Then
                            If fS Then
                                For i = 0 To txtTCS01.Count - 1
                                    dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKb11.Text)
                                    dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKb11.Text)
                                    txtCSB11(i).Text = CalcolaScarto(txtSB11.Text, dPrezzoCrr)
                                    txtCSSB11(i).Text = CalcolaScarto(txtSB11.Text, dPrezzoStd)
                                Next i
                                
                            End If
                            If fM Then
                                For i = 0 To txtTCM01.Count - 1
                                    dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKb11.Text)
                                    dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKb11.Text)
                                    txtCSB11(i).Text = CalcolaScarto(txtSB11.Text, dPrezzoCrr)
                                    txtCSSB11(i).Text = CalcolaScarto(txtSB11.Text, dPrezzoStd)
                                Next i
                                
                            End If
                            If fA Then
                                For i = 0 To txtTCA10.Count - 1
                                    dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKb11.Text)
                                    dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKb11.Text)
                                    txtCSB11(i).Text = CalcolaScarto(txtSB11.Text, dPrezzoCrr)
                                    txtCSSB11(i).Text = CalcolaScarto(txtSB11.Text, dPrezzoStd)
                                Next i
                                
                            End If
                        Else
                            If f1 Then
                                dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKA10.Text)
                                dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKA10.Text)
                                txtCSB11(0).Text = CalcolaScarto(txtSB11.Text, dPrezzoCrr)
                                txtCSSB11(0).Text = CalcolaScarto(txtSB11.Text, dPrezzoStd)
                            End If
                            If f2 Then
                                dPrezzoCrr = CDbl(txtTCC2.Text) ' + CDbl(txtEKb11.Text)
                                dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKb11.Text)
                                txtCSB11(1).Text = CalcolaScarto(txtSB11.Text, dPrezzoCrr)
                                txtCSSB11(1).Text = CalcolaScarto(txtSB11.Text, dPrezzoStd)
                            End If
                            If f3 Then
                                dPrezzoCrr = CDbl(txtTCC3.Text) ' + CDbl(txtEKb11.Text)
                                dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKb11.Text)
                                txtCSB11(2).Text = CalcolaScarto(txtSB11.Text, dPrezzoCrr)
                                txtCSSB11(2).Text = CalcolaScarto(txtSB11.Text, dPrezzoStd)
                            End If
                            If f5 Then
                                dPrezzoCrr = CDbl(txtTCC5.Text) ' + CDbl(txtEKb11.Text)
                                dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKb11.Text)
                                txtCSB11(3).Text = CalcolaScarto(txtSB11.Text, dPrezzoCrr)
                                txtCSSB11(3).Text = CalcolaScarto(txtSB11.Text, dPrezzoStd)
                            End If
                        End If
                Else
                    txtSB11.Text = "0"
                    For i = 0 To txtCSSB11.Count - 1
                        txtCSSB11(i).Text = "0"
                        txtCSB11(i).Text = "0"
                    Next i
                End If
                
                If Not IsNull(rsCiclo("Recupero")) And rsCiclo("Recupero") <> "" Or txtRB11.Text <> sRecupero Then
                    txtRB11.Text = rsCiclo("Recupero")
                    txtRCB11.Text = rsCiclo("CRecupero")
                For i = 0 To txtCRSB11.Count - 1
                    txtCRSB11(i).Text = CDbl((txtSB11.Text) / 100) * CDbl((txtRB11.Text / 100)) * CDbl(txtRCB11.Text)
                    txtCRCB11(i).Text = CDbl((txtSB11.Text) / 100) * CDbl((txtRB11.Text / 100)) * CDbl(txtRCB11.Text)
                Next i
                Else
                    txtRB11.Text = "0"
                    For i = 0 To txtCRSB11.Count - 1
                        txtCRSB11(i).Text = "0"
                        txtCRCB11(i).Text = "0"
                    Next i
                End If
'   Totali
                
                If fS Or fA Or fM Then
                    If fS Then
                        For i = 0 To txtTCS01.Count - 1
                            dPrezzoCrr = CDbl(txtTCS01(i).Text) ' + CDbl(txtEKb11.Text)
                            dPrezzoStd = CDbl(txtTSS01(i).Text) ' + CDbl(txtEKb11.Text)
                            txtTCB11(i).Text = CDbl(txtCSB11(i).Text) - CDbl(txtCRCB11(i).Text) + dPrezzoCrr + CDbl(txtEKB11.Text)
                            txtTSB11(i).Text = CDbl(txtCSSB11(i).Text) - CDbl(txtCRSB11(i).Text) + dPrezzoStd + CDbl(txtEKB11.Text)
                            Posizione txtCRSB11(i), True, txtCRSS01(i).Left, txtCRSB11(0).Top
                            Posizione txtCRCB11(i), True, txtCRCS01(i).Left, txtCRCB11(0).Top
                            Posizione txtCSB11(i), True, txtCSS01(i).Left, txtCSB11(0).Top
                            Posizione txtCSSB11(i), True, txtCSSS01(i).Left, txtCSSB11(0).Top
                            Posizione txtTCB11(i), True, txtTCS01(i).Left, txtTCB11(0).Top
                            Posizione txtTSB11(i), True, txtTSS01(i).Left, txtTSB11(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fM Then
                        For i = 0 To txtTCM01.Count - 1
                            dPrezzoCrr = CDbl(txtTCM01(i).Text) ' + CDbl(txtEKb11.Text)
                            dPrezzoStd = CDbl(txtTSM01(i).Text) ' + CDbl(txtEKb11.Text)
                            txtTCB11(i).Text = CDbl(txtCSB11(i).Text) - CDbl(txtCRCB11(i).Text) + dPrezzoCrr + CDbl(txtEKB11.Text)
                            txtTSB11(i).Text = CDbl(txtCSSB11(i).Text) - CDbl(txtCRSB11(i).Text) + dPrezzoStd + CDbl(txtEKB11.Text)
                            Posizione txtCRSB11(i), True, txtCRSM01(i).Left, txtCRSB11(0).Top
                            Posizione txtCRCB11(i), True, txtCRCM01(i).Left, txtCRCB11(0).Top
                            Posizione txtCSB11(i), True, txtCSM01(i).Left, txtCSB11(0).Top
                            Posizione txtCSSB11(i), True, txtCSSM01(i).Left, txtCSSB11(0).Top
                            Posizione txtTCB11(i), True, txtTCM01(i).Left, txtTCB11(0).Top
                            Posizione txtTSB11(i), True, txtTSM01(i).Left, txtTSB11(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
                    If fA Then
                        For i = 0 To txtTCA10.Count - 1
                            dPrezzoCrr = CDbl(txtTCA10(i).Text) ' + CDbl(txtEKb11.Text)
                            dPrezzoStd = CDbl(txtTSA10(i).Text) ' + CDbl(txtEKb11.Text)
                            txtTCB11(i).Text = CDbl(txtCSB11(i).Text) - CDbl(txtCRCB11(i).Text) + dPrezzoCrr + CDbl(txtEKB11.Text)
                            txtTSB11(i).Text = CDbl(txtCSSB11(i).Text) - CDbl(txtCRSB11(i).Text) + dPrezzoStd + CDbl(txtEKB11.Text)
                            Posizione txtCRSB11(i), True, txtCRSA10(i).Left, txtCRSB11(0).Top
                            Posizione txtCRCB11(i), True, txtCRCA10(i).Left, txtCRCB11(0).Top
                            Posizione txtCSB11(i), True, txtCSA10(i).Left, txtCSB11(0).Top
                            Posizione txtCSSB11(i), True, txtCSSA10(i).Left, txtCSSB11(0).Top
                            Posizione txtTCB11(i), True, txtTCA10(i).Left, txtTCB11(0).Top
                            Posizione txtTSB11(i), True, txtTSA10(i).Left, txtTSB11(0).Top
                            Select Case i
                                Case 1
                                    shpC2.Visible = True
                                Case 2
                                    shpC3.Visible = True
                                Case 3
                                    shpC5.Visible = True
                            End Select
                        Next i
                    End If
' totali senza lavorazioni intermedie
                Else
                    If f1 Then
                        dPrezzoCrr = CDbl(txtTCC1.Text) ' + CDbl(txtEKS01.Text)
                        dPrezzoStd = CDbl(txtTSC1.Text) ' + CDbl(txtEKS01.Text)
                        txtTCB11(0).Text = CDbl(txtCSB11(0).Text) - CDbl(txtCRCB11(0).Text) + dPrezzoCrr + CDbl(txtEKB11.Text)
                        txtTSB11(0).Text = CDbl(txtCSSB11(0).Text) - CDbl(txtCRSB11(0).Text) + dPrezzoStd + CDbl(txtEKB11.Text)
                        txtCRSB11(0).Visible = True
                        txtCRCB11(0).Visible = True
                        txtCSB11(0).Visible = True
                        txtCSSB11(0).Visible = True
                        txtTCB11(0).Visible = True
                        txtTSB11(0).Visible = True
                    End If
                    If f2 Then
                        dPrezzoCrr = CDbl(txtTCC2.Text) ' + CDbl(txtEKb11.Text)
                        dPrezzoStd = CDbl(txtTSC2.Text) ' + CDbl(txtEKb11.Text)
                        txtTCB11(1).Text = CDbl(txtCSB11(1).Text) - CDbl(txtCRCB11(1).Text) + dPrezzoCrr + CDbl(txtEKB11.Text)
                        txtTSB11(1).Text = CDbl(txtCSSB11(1).Text) - CDbl(txtCRSB11(1).Text) + dPrezzoStd + CDbl(txtEKB11.Text)
                        Posizione txtCRSB11(1), True, txtCRSC2.Left, txtCRSB11(0).Top
                        Posizione txtCRCB11(1), True, txtCRCC2.Left, txtCRCB11(0).Top
                        Posizione txtCSB11(1), True, txtCSC2.Left, txtCSB11(0).Top
                        Posizione txtCSSB11(1), True, txtCSSC2.Left, txtCSSB11(0).Top
                        Posizione txtTCB11(1), True, txtTCC2.Left, txtTCB11(0).Top
                        Posizione txtTSB11(1), True, txtTSC2.Left, txtTSB11(0).Top
                        shpC2.Visible = True
                    End If
                    If f3 Then
                        dPrezzoCrr = CDbl(txtTCC3.Text) ' + CDbl(txtEKb11.Text)
                        dPrezzoStd = CDbl(txtTSC3.Text) ' + CDbl(txtEKb11.Text)
                        txtTCB11(2).Text = CDbl(txtCSB11(2).Text) - CDbl(txtCRCB11(2).Text) + dPrezzoCrr + CDbl(txtEKB11.Text)
                        txtTSB11(2).Text = CDbl(txtCSSB11(2).Text) - CDbl(txtCRSB11(2).Text) + dPrezzoStd + CDbl(txtEKB11.Text)
                        Posizione txtCRSB11(2), True, txtCRSC3.Left, txtCRSB11(0).Top
                        Posizione txtCRCB11(2), True, txtCRCC3.Left, txtCRCB11(0).Top
                        Posizione txtCSB11(2), True, txtCSC3.Left, txtCSB11(0).Top
                        Posizione txtCSSB11(2), True, txtCSSC3.Left, txtCSSB11(0).Top
                        Posizione txtTCB11(2), True, txtTCC3.Left, txtTCB11(0).Top
                        Posizione txtTSB11(2), True, txtTSC3.Left, txtTSB11(0).Top
                        shpC3.Visible = True
                    End If
                    If f5 Then
                        dPrezzoCrr = CDbl(txtTCC5.Text) ' + CDbl(txtEKb11.Text)
                        dPrezzoStd = CDbl(txtTSC5.Text) ' + CDbl(txtEKb11.Text)
                        txtTCB11(3).Text = CDbl(txtCSB11(3).Text) - CDbl(txtCRCB11(3).Text) + dPrezzoCrr + CDbl(txtEKB11.Text)
                        txtTSB11(3).Text = CDbl(txtCSSB11(3).Text) - CDbl(txtCRSB11(3).Text) + dPrezzoStd + CDbl(txtEKB11.Text)
                        Posizione txtCRSB11(3), True, txtCRSC5.Left, txtCRSB11(0).Top
                        Posizione txtCRCB11(3), True, txtCRCC5.Left, txtCRCB11(0).Top
                        Posizione txtCSB11(3), True, txtCSC5.Left, txtCSB11(0).Top
                        Posizione txtCSSB11(3), True, txtCSSC5.Left, txtCSSB11(0).Top
                        Posizione txtTCB11(3), True, txtTCC5.Left, txtTCB11(0).Top
                        Posizione txtTSB11(3), True, txtTSC5.Left, txtTSB11(0).Top
                        shpC5.Visible = True
                    End If
                End If
End Sub

Private Function CalcolaMa514() As Double
Dim dDelta As Double, dStd As Double, dCrr As Double
Dim dMa514 As Double, dCosto As Double
dStd = txtCStMa514.Text
dCrr = txtCCMa514.Text
If Not IsNull(txtMescola(6).Text) And txtMescola(6).Text <> "" Then
    dMa514 = txtMescola(6).Text
Else
    MsgBox "Non ci sono percentuali di Ma514 in questa mescola!"
    Exit Function
End If

dDelta = dCrr - dStd
dCosto = dDelta * dMa514

CalcolaMa514 = dCosto


End Function
Private Function CalcolaAccoppiati()
Dim dPE As Double, dPET As Double, dT As Double
Dim dF As Double
    
    dPE = CDbl(txtMPS.Text) * Val(txtMPPS.Text)
    dPET = CDbl(txtSP.Text) * 1.33
    dF = dPE + dPET
    dT = CDbl(txtMPC.Text) * dF / 1000
'    dT = dT / (dF * 0.0354)

CalcolaAccoppiati = dT
End Function

Private Sub txtSP_DblClick()
txtSP.Locked = False
Me.Status = cMod

End Sub
