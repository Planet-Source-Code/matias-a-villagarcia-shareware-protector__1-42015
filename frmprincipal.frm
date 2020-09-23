VERSION 5.00
Begin VB.Form frmprincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial Generator"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2970
   Icon            =   "frmprincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   2970
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnumeroserial 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "&Save to file..."
      Height          =   495
      Left            =   1680
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdchequear 
      Caption         =   "&Check"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "&Generate"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtnombre 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtnumero4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtnumero3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtnumero2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtnumero1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   2
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblnumeroserial 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Line Line3 
      X1              =   2160
      X2              =   2280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   1560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   840
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblserial 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label lblnombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmprincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE: The serial number is case-sensitive
'You can use this code in your
'programs freely but put me in the credits please, and if
'you can let me know.
'Thanks, Matías Ariel Villagarcía.

Private Sub cmdchequear_Click()
If Check(txtnombre.Text, txtnumero1.Text & "-" & txtnumero2.Text & "-" & txtnumero3.Text & "-" & txtnumero4.Text, Mid(App.Path, 1, 3)) = True Then
    MsgBox "Correct Serial Number.", vbInformation, "Serial Number"
Else
    MsgBox "Wrong Serial Number.", vbExclamation, "Serial Number"
End If
End Sub

Private Sub cmdgenerar_Click()
txtnumeroserial.Text = GenerateSerialHD(txtnombre.Text, Mid(App.Path, 1, 3))
End Sub

Private Sub cmdguardar_Click()
Open App.Path & "\" & "Serials.txt" For Append As #1
    Print #1, "Name: " & txtnombre.Text & " Serial Number: " & txtnumero1.Text & "-" & txtnumero2.Text & "-" & txtnumero3.Text & "-" & txtnumero4.Text
Close #1
MsgBox "Data successfully saved to disk.", vbInformation, "File Saved"
End Sub

Private Sub cmdsalir_Click()
End
End Sub

Private Sub Form_Load()
frmprincipal.Caption = "Serial Generator Version: " & App.Major & "." & App.Minor & "." & App.Revision & " Por Matías A. V."
End Sub

Private Sub txtnombre_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdgenerar_Click
    txtnumero1.SetFocus
End If
End Sub

Private Sub txtnumero1_Change()
If Len(txtnumero1.Text) = 4 Then txtnumero2.SetFocus
End Sub

Private Sub txtnumero1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnumero2_Change()
If Len(txtnumero2.Text) = 4 Then txtnumero3.SetFocus
If Len(txtnumero2.Text) = 0 Then txtnumero1.SetFocus
End Sub

Private Sub txtnumero2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnumero3_Change()
If Len(txtnumero3.Text) = 4 Then txtnumero4.SetFocus
If Len(txtnumero3.Text) = 0 Then txtnumero2.SetFocus
End Sub

Private Sub txtnumero3_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnumero4_Change()
If Len(txtnumero4.Text) = 4 Then cmdchequear.SetFocus
If Len(txtnumero4.Text) = 0 Then txtnumero3.SetFocus
End Sub

Private Sub txtnumero4_GotFocus()
SendKeys "{Home}+{End}"
End Sub
