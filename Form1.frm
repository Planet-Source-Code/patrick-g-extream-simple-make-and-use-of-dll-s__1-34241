VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "&About"
      Height          =   255
      Left            =   4080
      MaskColor       =   &H8000000D&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "It is just a very basic example of what you can do with a DLL.  If enough votes come in i'll make this 99x better."
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Uptime:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filesize:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dir Exists:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer name:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Exists:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "App Version:"
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IF YOU GOT AN ERROR ON STARTUP DON'T WORRY
'DO THE FOLLOWING THAT I TELL YOU TO DO BELOW AND IT WILL WORK!!!
'
'I will be adding many many more functions.
'Only if i see that people like my idea and find it can be useful
'I am trying to win on PSC one last time,
'but i do not want to win unless i deserve it!
'So let me know how my DLL can help you! =)
'
'To load the DLL in to your project do the following!
'
'Goto Project Menu > Refrences then hit browse
'Look for the MultiFuncDLL.dll file on your computer and hit ok!
'Then hit OK again
'
'Now in your project i'll use a "Form Load EG"
'Dim the following into your project
'
'dim Dll as multifuncdll.extras
'
'public sub form_load()
'set dll = new extras
'
'From there you can call any function like this
'dll.osuptime (And whatever else is in the selection menu
'
'end sub

Dim dll As ExtrasDLL.Extras

Private Sub Command1_Click()
Call dll.Show
End Sub

Public Sub form_load()
'This loads the dll to be used
Set dll = New Extras
'now we will place all items in the text boxes as an example
Text1.Item(0) = dll.AppVersion
If dll.FileCheck("C:\windows\command.com") = True Then
    Text1.Item(1) = "Command.com Exists"
Else
    Text1.Item(1) = "Command.com Doesn't Exist"
End If

Text1.Item(2) = dll.Computername

If dll.FileDirCheck("C:\windows") = True Then
    Text1.Item(3) = "Windows Dir Exists"
Else
    Text1.Item(3) = "Windows Dir Does Not Exist"
End If

Text1.Item(4) = dll.FileSize("C:\windows\system32\diskcopy.dll")
Text1.Item(5) = dll.OsUptime

End Sub

Private Sub Form_Unload(Cancel As Integer)
'This is how simple you can get an exit message
Call dll.ExitMessage("Testing", "Test", "Ok", "No", True)
End Sub
