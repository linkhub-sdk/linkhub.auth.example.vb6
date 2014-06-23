VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "링크허브 인증 API 예제"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton btnGetPartnerBalance 
      Caption         =   "파트너 잔여 통합포인트 확인"
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   1680
      Width           =   3135
   End
   Begin VB.CommandButton btnGetBalance 
      Caption         =   "회원 잔여포인트 확인"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtSession_Token 
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   6615
   End
   Begin VB.CommandButton btnGetToken 
      Caption         =   "토큰 발급"
      Height          =   1335
      Left            =   6960
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLinkhub As Linkhub

Private Const serviceID = "POPBILL_TEST"
Private Const linkID = "TESTER"
Private Const SecretKey = "088b1258aoeMH5OtGjK4zaOlwZGVvSK40ceI8t4j7Hw="


Private Sub btnGetBalance_Click()
    Dim remainPoint As Double
    
    remainPoint = mLinkhub.GetBalance(txtSession_Token.Text, serviceID)
    
    If remainPoint < 0 Then
        MsgBox ("[" + CStr(mLinkhub.LastErrCode) + "] " + mLinkhub.LastErrMessage)
        Exit Sub
    End If
    
    
    MsgBox "잔여포인트 : " + CStr(remainPoint)
    
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim remainPoint As Double
    
    remainPoint = mLinkhub.GetPartnerBalance(txtSession_Token.Text, serviceID)
    
    If remainPoint < 0 Then
        MsgBox ("[" + CStr(mLinkhub.LastErrCode) + "] " + mLinkhub.LastErrMessage)
        Exit Sub
    End If
    
    
    MsgBox "파트너 잔여포인트 : " + CStr(remainPoint)
End Sub

Private Sub btnGetToken_Click()
    Dim token As LinkhubToken
    
    Dim scope As New Collection
     
    scope.Add "member"
    scope.Add "110"
    
    Set token = mLinkhub.getToken(serviceID, "1231212312", scope)
    
    If token Is Nothing Then
        MsgBox ("[" + CStr(mLinkhub.LastErrCode) + "] " + mLinkhub.LastErrMessage)
        Exit Sub
    End If
    
    txtSession_Token.Text = token.session_token
    
    
End Sub

Private Sub Form_Load()
  
    Set mLinkhub = New Linkhub
    
    mLinkhub.linkID = linkID
    mLinkhub.SercetKey = SecretKey
    mLinkhub.IsTest = True
    
End Sub

