VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Whiteboard 
   Caption         =   "Internet Whiteboard"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   8745
   Begin VB.PictureBox picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   -960
      ScaleHeight     =   4995
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   360
      Width           =   7815
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   7560
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Whiteboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Option Explicit
'*Begin Discription
'This code creates a simple whiteboard that creates graphics on
'both picture boxes by passing the x and y coordintes across the
'TCPIP connection. Regardless of which person draws the picture
'remains the same.
'The app uses one form, one picture box and a winsock control
'*End Discription
'*Begin Keywords
'*End Keywords
'*Begin Topics
'*End Topics
Dim CellHeight As Single, CellWidth As Single, Cells() As String, Turn As Integer
Dim CurrentMark As String
Private pbListener As Boolean


Private Sub Form_Load()
    'Check for TCPIP connection. This check will set up a second
    'connection to the first accross the tcpip and ports if
    'a TCPIP connection is available
    Dim sCaption As String
    Dim lInstances As Long
    
    'the port is part of the caption, trim it
    Me.Caption = "Internet Whiteboard"
    'run only two instances
   sCaption = Me.Caption
   lInstances = AppInstances(sCaption)
   pbListener = lInstances = 1
   'prevent more than 2 instances
   If lInstances > 2 Then End
    
 
    
    If Len(Winsock1(0).LocalIP) > 0 Then
        'For this sample, since the sender and receiver are the
        'same, the RemoteHost(TCPIP) is set to the localIP
        
        Winsock1(0).RemoteHost = Winsock1(0).LocalIP
        'If this is the first instance of the program
        
        If pbListener Then
          
            'On the same machine the remote port must be
            'different than the binding port
          
            
     
            Winsock1(0).LocalPort = 101
            Winsock1(0).Listen
            'Position at the top left corner
            Whiteboard.Left = 0
            Whiteboard.Top = 0
            Whiteboard.Width = Screen.Width / 2
            picture1.Width = Whiteboard.Width - 400
            'To Simulate second app Start a  new instance of
            'the program
            Shell App.Path & "\WhiteBrd.exe", vbNormalNoFocus
            'The second instance of the program
        Else
            
 

            Winsock1(0).RemotePort = 101
    
            Winsock1(0).RemoteHost = Winsock1(0).LocalIP
            Winsock1(0).Connect
           
            'Center width of screen to the right
            Whiteboard.Left = Screen.Width / 2
            Whiteboard.Top = 0
            Whiteboard.Width = Screen.Width / 2
            picture1.Width = Whiteboard.Width - 400
        End If
    End If
    'Display the current port in the caption
 Me.Caption = Me.Caption & " Local Port - " & Winsock1(0).LocalPort
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Winsock1(0).State <> sckClosed Then Winsock1(0).Close

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If left mouse button is pessed
    If Button = 1 Then
        picture1.PSet (X, Y)
        'if connected send x and y to bound program
        Winsock1(0).SendData X & "," & Y
    End If
End Sub

Private Sub Winsock1_Close(Index As Integer)
Unload Me
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
If Index = 0 And pbListener Then
    If Winsock1(Index).State <> sckClosed Then Winsock1(Index).Close
    Winsock1(Index).Accept requestID
End If

End Sub

'This procedure processes data that arrives across the TCPIP Connection
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim SData As String, SX As Single, SY As Single
    'When data arrives if the window is minimized maximize it
    If Whiteboard.WindowState = 1 Then Whiteboard.WindowState = 0
    'Put the data into a string
    Winsock1(0).GetData SData
    'Extract the X
    SX = Mid(SData, 1, InStr(1, SData, ",") - 1)
    'and Y values
    SY = Mid(SData, InStr(1, SData, ",") + 1)
    'Create the graphic
    picture1.PSet (SX, SY)
End Sub
