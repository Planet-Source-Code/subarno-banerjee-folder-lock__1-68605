VERSION 5.00
Begin VB.Form frmL 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock Folder"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "frmL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3795
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Type Your PassWord"
      Top             =   570
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Type your User Name"
      Top             =   83
      Width           =   2175
   End
   Begin VB.CommandButton cmdLock 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lock"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      MouseIcon       =   "frmL.frx":AB7A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Lock Folder"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Do Not Forget This Name and PassWord."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lnlPW 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   600
      Width           =   1230
   End
   Begin VB.Label lblNm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Left            =   510
      TabIndex        =   3
      Top             =   120
      Width           =   705
   End
End
Attribute VB_Name = "frmL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLock_Click()
If MsgBox("Do you really want to Lock the folder ?", vbYesNo, "Lock ??") = vbYes Then
frmMain.Nm = txtName.Text
frmMain.Pw = txtPW.Text
Unload Me
Else
txtName.Text = ""
txtPW.Text = ""
txtName.SetFocus
End If
End Sub
Private Sub txtName_Change()
cmdLock.Enabled = CBool(Len(txtName.Text)) And CBool(Len(txtPW.Text))
End Sub
Private Sub txtPW_Change()
Call txtName_Change
End Sub
Private Sub txtPW_KeyPress(KeyAscii As Integer)
If txtName.Text <> "" And txtPW.Text <> "" And KeyAscii = 13 Then
Call cmdLock_Click
End If
End Sub
