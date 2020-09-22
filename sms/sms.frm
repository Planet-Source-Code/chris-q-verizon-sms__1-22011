VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Verizon SMS                     by CQVB"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3165
   Icon            =   "sms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Character count"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   1455
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   4320
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4560
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Height          =   495
      Left            =   2040
      Picture         =   "sms.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdSend 
      Height          =   495
      Left            =   120
      Picture         =   "sms.frx":1EF4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Number"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         MaxLength       =   10
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Message - max 120"
      ForeColor       =   &H0000FF00&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
      Begin VB.TextBox Text2 
         Height          =   1095
         Left            =   120
         MaxLength       =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   0
      Picture         =   "sms.frx":4DE6
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
sendmessage                 'performs sendmessage function
Text1.Text = ""             'clears number
Text2.Text = ""             'clears message
End Sub
Private Sub cmdExit_Click()
End                         'closes program
End Sub
Function sendmessage()

Dim strServlet As String, strMsgData As String
    'this is the url to the servlet, to see the online form
    'goto  http://www.msg.myvzw.com/messaging.jsp
    strServlet = "http://www.msg.myvzw.com/servlet/SmsServlet"
    'view source of the online form and you will see the text inputs
    'that are named 'min' and 'message', these are the only two inputs
    'required
    
    
    'this is the number and message--Note: you could actually send the message
    'to 10 different numbers simultaneously by separating the numbers
    'with commas, but for demonstrations purposes, i've set the maxlength to 10
    'in the Number(text1) box.
    strMsgData = "&min=" & Text1.Text & "&message=" & Text2.Text
    
    Inet1.Execute strServlet, "Post", strMsgData, _
    "Content-Type: application/x-www-form-urlencoded"
End Function


Private Sub Form_Load()
Timer1.Enabled = True       'starts timer for character counting

End Sub

'keeps count of characters in message box
Private Sub Timer1_Timer()
numcharacters = Len(Text2.Text)
Label1.Caption = numcharacters
End Sub
