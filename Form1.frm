VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RC4 Encryption Module Demo"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt/Decrypt File"
      Default         =   -1  'True
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pass Phrase"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Password"
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output File:"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File:"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "This will run much slower in the VB design environment."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
  Dim FileName As String
  Dim FileNum As Integer
  Dim FileBytes() As Byte

  'Open File Here
  'Veriable = "FileName"
  
  FileName = text2
  
  'Reading File
  ReDim FileBytes(FileLen(FileName) - 1)
  FileNum = FreeFile
  Open FileName For Binary Access Read As FileNum
    Get #FileNum, , FileBytes
  Close FileNum

  'Encrypting/Decrypting Data
  RC4 FileBytes, Text1.Text 'Password

  'Save File Here
  'Veriable = "FileName"
  FileName = text3

  'Writing File
  FileNum = FreeFile
  Open FileName For Binary Access Write As FileNum
    Put #FileNum, , FileBytes
  Close FileNum
End Sub
