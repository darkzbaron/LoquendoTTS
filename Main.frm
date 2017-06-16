VERSION 5.00
Object = "{5A80A7F3-4AC0-44B4-B684-1F18ADC68C4D}#1.0#0"; "LTTS7.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loquendo TTS Sample"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin LTTS7Lib.LTTS7 objLTTS 
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.CommandButton btnRecord 
      Caption         =   "Record to &file"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton btnPause 
      Caption         =   "&Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   1695
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "&Language:"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "&Voice:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "&Insert text:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lbEvents_Click()

End Sub


Private Sub btnRecord_Click()
    Static sFileName, sLastFileName As String
    
    If Trim(txtInput) = "" Then
        MsgBox "No text specified!"
        Exit Sub
    End If
    If sLastFileName = "" Then sLastFileName = "c:\Sample.wav"
    sFileName = InputBox("Please enter a full file pathname:", , sLastFileName)
    If Trim(sFileName) = "" Then
        MsgBox "No name specified: record aborted!"
        Exit Sub
    End If
    sLastFileName = sFileName
    
    MousePointer = vbHourglass
   
    If Not objLTTS.Record(txtInput, sFileName) Then
        MousePointer = vbDefault
        MsgBox "Error while recording to file!"
        Exit Sub
    End If
  
    MousePointer = vbDefault
End Sub
Sub Exit_Form()

MsgBox ("ok")
End Sub
Sub Form_Load()
    Dim s As String
    Dim m As String
    Dim txtInput As String
    Dim sampleString As String
    'm = objLTTS.Left
    txtInput = Command
    objLTTS.Language = "EnglishUs"
    objLTTS.Voice = "Allison"
    objLTTS.AudioChannels = "Stereo"
    objLTTS.Frequency = "48000"
        
    'txtInput = "\item=Laugh"
    objLTTS.Read txtInput
    
    frmMain.Hide

        
   
    
    'Gosub Exit _Form()
  
   
    
    

End Sub
Private Sub objLTTS_EndOfSpeech()
    
   End
End Sub


