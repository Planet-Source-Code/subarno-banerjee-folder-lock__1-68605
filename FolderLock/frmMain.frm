VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lock / Unlock Folder"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   3375
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnLk 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Unlock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      MouseIcon       =   "frmMain.frx":AB7A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Unlock Selected Folder"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdLk 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Lock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2160
      MouseIcon       =   "frmMain.frx":BFC4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Lock Selected Folder"
      Top             =   840
      Width           =   1095
   End
   Begin VB.DirListBox dir 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1830
   End
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
   Begin VB.Menu mnuLst 
      Caption         =   "Help"
      Begin VB.Menu mnuIns 
         Caption         =   "Instructions"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAb 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const AlterExtn = ".{f39a0dc0-9cc8-11d0-a599-00c04fd64433}" 'Class ID for Channel File
Public Nm, Pw As String
Dim FSO As New FileSystemObject
Dim MyDir
Private Sub cmdLk_Click()
On Error GoTo Err_Rep
If InStr(1, MyDir, ".", vbTextCompare) = 0 Then
frmL.Show vbModal, frmMain
If Nm <> "" And Pw <> "" Then
Open MyDir & "\Protect.dat" For Output As #1
Print #1, Nm
Print #1, Pw
Close #1
Name MyDir As MyDir & AlterExtn
Call MsgBox("The folder has been locked." & vbCrLf & vbCrLf & "You can Unlock it anytime by providing the same Name and PassWord.", vbInformation, "Folder Locked")
Clear
End If
Else
Call MsgBox("This folder is already Locked by other users.", vbExclamation, "Access Denied")
End If
Err_Rep:
If Err Then
Call MsgBox(Err.Description, vbCritical, "Error !")
End If
End Sub
Private Sub cmdUnLk_Click()
On Error GoTo Err_Rep
If InStr(1, MyDir, ".", vbTextCompare) > 0 Then
Dim TempNm, TempPw As String
frmUnL.Show vbModal, frmMain
If Nm <> "" And Pw <> "" Then
Open MyDir & "\Protect.dat" For Input As #2
Line Input #2, TempNm
Line Input #2, TempPw
Close #2
If TempNm <> Nm Or TempPw <> Pw Then
Call MsgBox("You are not authorised to Unlock this folder." & vbCrLf & vbCrLf & "This folder has been Locked by other users and can be Unlocked exclusively by the Owner user.", vbExclamation, "Unauthorised intrusion not allowed.")
Exit Sub
End If
Kill MyDir & "\Protect.dat"
Name MyDir As Mid(MyDir, 1, InStr(1, MyDir, ".", vbTextCompare) - 1)
Call MsgBox("The folder has been unlocked." & vbCrLf & vbCrLf & "You can now access it.", vbInformation, "Folder Unlocked")
Clear
End If
Else
Call MsgBox("This folder isn't Locked at all. You are free to access it.", vbExclamation, "Access Denied")
End If
Err_Rep:
If Err Then
Call MsgBox(Err.Description, vbCritical, "Error !")
End If
End Sub
Private Sub dir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MyDir = dir.List(dir.ListIndex)
End Sub
Private Sub drv_Change()
On Error GoTo Err_Rep
dir.Path = drv.Drive
Err_Rep:
If Err Then
Call MsgBox(Err.Description, vbCritical, "Error !")
End If
End Sub
Private Sub Form_Load()
dir.Path = FSO.GetParentFolderName(dir.Path)
End Sub
Private Sub Clear()
Nm = ""
Pw = ""
dir.Refresh
End Sub
Private Sub mnuAb_Click()
Call MsgBox("This Application has been designed by -" & vbCrLf & "               Subarno Banerjee." & vbCrLf & "Inspired by the motivation of his friend, Subhasish Hazra." & vbCrLf & vbCrLf & "This application can be very useful in protecting personal or confidential files. Besides, it also helps those who have to share their resources like that in a cyber-cafe or school. You are free to use, comment on, suggest and modify this code.", vbInformation, "About")
End Sub
Private Sub mnuIns_Click()
Call MsgBox("Select the desired folder in the Directory List Box and click on the 'Lock' or 'Unlock' button to lock or unlock the folder as desired. Feed your User Name and Password correctly and your job is done. Folders protected by this application won't open in Windows Explorer unless you Unlock it.", vbInformation, "Instructions")
End Sub
