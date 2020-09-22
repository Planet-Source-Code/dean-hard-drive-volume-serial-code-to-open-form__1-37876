VERSION 5.00
Begin VB.Form frmSerial 
   ClientHeight    =   1560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   1560
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2595
   End
   Begin VB.TextBox txtDisk 
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "C DRIVE IS DEFAULT"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Serial No of this Drive is"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1875
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Disk As String
Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_Load()
' It will give the serial no of the Drive written in textbox
Disk = txtDisk.Text
txtSerialNo.Text = VolumeSerialNumber(Disk)
If txtSerialNo.Text = "317F-0AF3" Then
frmSerial.Hide
Form1.Show
Else: MsgBox "YOUR KEYCODE IS NOT CORRECT ASK VENDOR FOR KEYCODE", 16, "KEYCODE NOT CORRECT"
End
End If
End Sub
