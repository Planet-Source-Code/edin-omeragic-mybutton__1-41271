VERSION 5.00
Begin VB.Form InputBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MyButtonProject.MyButton btnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   3
      Top             =   720
      Width           =   1155
      _ExtentX        =   1931
      _ExtentY        =   661
      SPN             =   "MyButtonDefSkin"
      Text            =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MyButtonProject.MyButton btnOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   2
      Top             =   240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   661
      Text            =   "&Ok"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   300
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Type your name"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "InputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_Response As VbMsgBoxResult

Private Sub Form_Load()

    Dim B As MyButton
    Dim C As Control
    For Each C In Me.Controls
        If TypeName(C) = "MyButton" Then
            Set B = C
            Set B.SkinPicture = MyButtonDemo.MyButtonDefSkin
        End If
    Next


End Sub

Private Sub btnOK_Click()
    m_Response = vbOK
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    m_Response = vbCancel
    Me.Hide
End Sub

Public Property Get Response() As VbMsgBoxResult
    Response = m_Response
End Property

Public Property Get Text() As String
    Text = Text1.Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    Text1.Text = vNewValue
End Property
