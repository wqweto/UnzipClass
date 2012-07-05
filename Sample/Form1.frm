VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6744
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   8976
   LinkTopic       =   "Form1"
   ScaleHeight     =   6744
   ScaleWidth      =   8976
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5616
      Left            =   84
      TabIndex        =   1
      Top             =   84
      Width           =   8832
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   768
      Left            =   3360
      TabIndex        =   0
      Top             =   5880
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' $Header: $
'
'   VB6 Unzip Support
'   Copyright (c) 2012 Unicontsoft
'
'   Sample form
'
' $Log: $
'
'=========================================================================
Option Explicit

Private WithEvents m_oZip As cZip
Attribute m_oZip.VB_VarHelpID = -1

Private Sub Command1_Click()
    Dim lIdx            As Long
    Dim baFiles()       As Boolean
    Dim dblTimer        As Double
    
    dblTimer = Timer
    Set m_oZip = New cZip
    m_oZip.Init App.Path & "\Build_2011_11_30.zip"
'    m_oZip.Init "D:\TEMP\Dreem15_Meik_db_201203200251.zip"
'    m_oZip.Init App.Path & "\aaa.zip"
'    m_oZip.Init App.Path & "\Dreem15_Hubo_db_201203210147.zip"
    For lIdx = 0 To m_oZip.Count - 1
        List1.AddItem m_oZip.File(lIdx).FileName & vbTab & m_oZip.File(lIdx).CompressedSize & vbTab & m_oZip.File(lIdx).DateTime
    Next
    ReDim baFiles(0 To m_oZip.Count)
    For lIdx = 0 To m_oZip.Count - 1
        baFiles(lIdx) = True
    Next
    m_oZip.Unzip App.Path & "\test_unpack" ' , baFiles
    Caption = Round(Timer - dblTimer, 3)
End Sub

Private Sub m_oZip_Error(ByVal Idx As Long, Error As String)
    List1.AddItem m_oZip.File(Idx).FileName & " " & Error
End Sub

Private Sub m_oZip_Progress(ByVal Idx As Long, ByVal Pos As Long, ByVal Size As Long)
    Caption = m_oZip.File(Idx).FileName & " " & Format(Pos * 100# / Size, "0.0") & "%"
    DoEvents
End Sub
