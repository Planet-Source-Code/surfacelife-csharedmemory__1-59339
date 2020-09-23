VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5070
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List2 
      Height          =   8445
      Left            =   4710
      TabIndex        =   1
      Top             =   165
      Width           =   4275
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4155
      Top             =   2070
   End
   Begin VB.ListBox List1 
      Height          =   8445
      Left            =   195
      TabIndex        =   0
      Top             =   180
      Width           =   4275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sh As New CSharedMemory

Private Sub Form_Load()
    Winsock1.Protocol = sckUDPProtocol
'    Winsock1.RemoteHost = "255.255.255.255"
    Winsock1.LocalPort = 12345
    Winsock1.Bind 12345
    Dim i As Integer
    For i = 0 To 39
        List1.AddItem i + 1
    Next
    sh.sharedMemory "JOY", 80
    Timer1.Interval = 0
    Show
    Do
        DoEvents
        Timer1_Timer
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
        ' 0 - js1 buffer index
        ' 1 - js2 buffer index
        ' 2-41 - js1 buffer 1
        ' 42-81 - js1 buffer 2
        ' 82-121 - js2 buffer 1
        ' 122-161- js2 buffer 2
    Dim i As Integer, s As String, b() As Byte, j As Integer
    b = sh.memoryb(0, 1)
    j = b(0)
    Caption = j
    b = sh.memoryb(2 + 40 * j, 40)
    For i = 0 To UBound(b)
        List1.List(i) = b(i)
    Next
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim i As Integer, s As String, b() As Byte, j As Integer
    Winsock1.GetData s
    'Winsock1.Bind Winsock1.LocalPort
    b = StrConv(s, vbFromUnicode)
    For i = 0 To UBound(b)
        List2.List(i) = b(i)
    Next
    
End Sub

