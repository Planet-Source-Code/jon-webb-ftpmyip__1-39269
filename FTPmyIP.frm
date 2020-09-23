VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP My IP Address"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4140
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   180
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   2355
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINNT\System32\shdoclc.dll/dnserror.htm#http:///"
   End
   Begin VB.Label lbldone 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private moFTP As cFTP
Private Sub FTPit()
Dim filename1, filename2 As String
filename1 = "C:\myip.txt"
filename2 = "myip.txt"

    moFTP.PutFile filename1, filename2, ftAscii
  
    lbldone.Caption = "File Uploaded @ " & Format(Now, "hh:mm:ss")

End Sub
Private Sub CONNECTit()
On Error GoTo vbErrorHandler
' Connect to server!'
    
    Screen.MousePointer = vbHourglass
' Setup the FTP Object
'
    With moFTP
        .Host = "ftp.DOMAIN.co.uk"
        .User = "USERNAME"
        .Password = "PASSWORD"
    End With
    
    moFTP.Connect

    Screen.MousePointer = vbDefault

    Exit Sub
vbErrorHandler:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description

End Sub

Private Sub Form_Load()
    Set moFTP = New cFTP
Dim msgtxt As String
'load my ip address
WebBrowser1.Navigate "http://www.facultyof1000.com/whatsmyip.asp"
'wait till it's loaded
Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE
    DoEvents
Loop
'get the text
msgtxt = WebBrowser1.Document.documentelement.innertext
    Open "c:\myip.txt" For Output As #1
    Print #1, msgtxt
    Close #1
CONNECTit
FTPit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set moFTP = Nothing
End Sub
