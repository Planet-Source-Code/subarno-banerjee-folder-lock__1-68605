VERSION 5.00
Begin VB.Form frmUnL 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unlock Folder"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmUnL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3945
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      ToolTipText     =   "Type your User Name"
      Top             =   180
      Width           =   2175
   End
   Begin VB.TextBox txtPW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Type Your PassWord"
      Top             =   660
      Width           =   2175
   End
   Begin VB.CommandButton cmdUnL 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Unlock"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      MouseIcon       =   "frmUnL.frx":144A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Unlock Folder"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblNm 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   585
      TabIndex        =   4
      Top             =   240
      Width           =   630
   End
   Begin VB.Label lblPW 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PassWord"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmUnL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUnL_Click()
If MsgBox("Do you really want to Unlock the folder ?", vbYesNo, "Unlock ??") = vbYes Then
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
cmdUnL.Enabled = CBool(Len(txtName.Text)) And CBool(Len(txtPW.Text))
End Sub
Private Sub txtPW_Change()
Call txtName_Change
End Sub
Private Sub txtPW_KeyPress(KeyAscii As Integer)
If txtName.Text <> "" And txtPW.Text <> "" And KeyAscii = 13 Then
Call cmdUnL_Click
End If
End Sub
