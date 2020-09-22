VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Long to Binary Conversion"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtBinaryValue 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox txtLongValue 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Type your (long) value here then press enter..."
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Enter a (long) value then press Enter to convert that value into binary."
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblBinaryValue 
      Caption         =   "Binary (string)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label lblLongValue 
      Caption         =   "Value (long)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   165
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CR1 = vbNewLine
Private Const CR2 = CR1 & CR1

Private Sub cmdExit_Click()
  Unload Me
End Sub

' This function will be called whether the user clicks the Exit button, or the [X] button
' on the top right of the frmMain...
Private Sub Form_Unload(Cancel As Integer)
  ' Note the setting of the default button to "No". Always good practice when
  ' asking this type of question
  If MsgBox("Exit this demonstration of recursion by" & CR1 & _
            "Dave@DRLDEV.CO.UK" & CR2 & _
            "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton2, "Exit") = vbNo Then
    Cancel = 1
  End If
End Sub

' Allows the user to enter values and just press the enter key to initiate the conversion
' without having to use the mouse
Private Sub txtLongValue_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then ' Enter has been pressed
    Me.txtBinaryValue = LongToBinary(Val(Me.txtLongValue)) ' convert to binary
    ' This is a trick to allow the user to type the new value without deleting the old
    ' one, done by highlighting the old value so that it is replaced when the user begins
    ' entering the new value...
    Me.txtLongValue.SelStart = 0
    Me.txtLongValue.SelLength = Len(Me.txtLongValue)
  End If
End Sub

