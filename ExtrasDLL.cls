VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Extras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Below i will explain how to use each function.
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

Option Explicit

Public Function MathAdd(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathadd("1", "2") will give you a result with "3"
'Useage: msgbox can be used as well along with everything else
    MathAdd = FirstNumber + SecondNumber
End Function

Public Function MathSubtract(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathsubtract("2", "1") will give you a result with "1"
'Useage: msgbox can be used as well along with everything else
    MathSubtract = FirstNumber - SecondNumber
End Function

Public Function MathDivide(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathdivide("2", "1") will give you a result with "1"
'Useage: msgbox can be used as well along with everything else
    MathDivide = FirstNumber / SecondNumber
End Function

Public Function MathMultiply(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
'Useage: text1.text = dll.mathMultiply("12", "12") will give you a result with "144"
'Useage: msgbox can be used as well along with everything else
    MathMultiply = FirstNumber * SecondNumber
End Function

Public Function RandomNum(ByVal PickaNumber As Long)
'Useage: text1.text = dll.RandomNum("Any number") will pick a random number between your value
'You can also just leave it blank and just use dll.RandomNum and it will pick a number between 8000
    If RandomNum = "" Then
        RandomNum = Int(Rnd * 8001)
    Else
        RandomNum = Int(Rnd * PickaNumber)
    End If
End Function

Public Function Computername() As String
'Useage: text1.text = dll.computername will return the network name of the computer
'Useage: msgbox can be used as well along with everything else
  Dim lpBuff   As String * 25
  Dim retval   As Long
  retval = GetComputername(lpBuff, 25)
  Computername = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function

Public Sub SaveToINI(ByVal vntValue As Variant, ByVal sSection As String, ByVal sKey As String, SFilename As String)
'Useage: text1.text = dll.savetoini("Keyname", "Section", text1.text, appdir & "test.ini")
    #If Win32 Then
        Dim xRet          As Long
    #Else
        Dim xRet          As Integer
    #End If
    
    xRet = WritePrivateProfileString(sSection, sKey, CStr(vntValue), SFilename)
End Sub
Public Function GetFromINI(ByVal sSection As String, ByVal sKey As String, SFilename As String) As Variant
'Useage: text1.text = dll.getfromini("Keyname", "Section", "String", appdir & "test.ini")
'Useage: msgbox can be used as well along with everything else
    #If Win32 Then
        Dim xRet          As Long
    #Else
        Dim xRet          As Integer
    #End If

Dim sReturnStr    As String
Dim nStringLen    As Integer

    nStringLen = 255
    sReturnStr = String(nStringLen, Chr$(0))
    xRet = GetPrivateProfileString(sSection, sKey, "", sReturnStr, nStringLen, SFilename)
    GetFromINI = Left(sReturnStr, xRet)
End Function

Public Function ExitMessage(ByVal Message As String, ByVal Title As String, ByVal IfYesMessage As String, ByVal IfNoMessage As String, Optional YesNo As Boolean)
'Info: Exit message is something that just saves time, isn't absolutely useful but sufficiant
'Useage: To get Yes or No message use,
'Useage: call dll.exitmessage("Enter Exit Message", "Msgbox Title", "IF answer is yes message", "If answer is no message", true)
'Useage: To show just a regular exit message use,
'Useage: call dll.exitmessage("Enter Exit Message", "Msgbox Title", "", "", false)
    If YesNo = True Then
        Dim Answer
        Answer = MsgBox(Message, vbYesNo Or vbInformation, Title)
            If Answer = vbYes Then
               MsgBox IfYesMessage, vbInformation, Title
            Else
                If IfNoMessage = "end" Then
                    Dim Form As Form
                    Unload Form
                    Set Form = Nothing
                Else
                    Dim Cancel As Integer
                    MsgBox IfNoMessage, vbInformation, Title
                    Cancel = 1
                    Exit Function
                End If
            End If
    Else
        MsgBox Message, vbInformation, Title
    End If
End Function

Public Function AppIsRunning(Message As String) As Boolean
'Useage: Use the follow for eg on how to use this
'
'if dll.appisrunning = true then
'   msgbox "App is already running", vbinformation, "App"
'end if
    If (App.PrevInstance = True) Then
        AppIsRunning = True
    Else
        AppIsRunning = False
    End If
End Function

Public Function AppVersion()
'Useage: very self explanitory!!!
'Msgbox dll.appversion
    AppVersion = App.Major & "." & App.Major & "." & App.Revision
End Function

Public Function FileCheck(Path$) As Boolean
'Info: Gives ability to check if a file exists or not
'Useage: call dll.filecheck("C:\windows\command.com")
    FileCheck = True
On Error Resume Next
    Dim Disregard As Long
    Disregard = FileLen(Path)
    If Err <> 0 Then FileCheck = False
End Function

Function OsExitWindows(BootMode As Integer)
'Info: Allows you to shutdown or reboot your computer with small little command
'Useage: dll.osexitwindows("1") Shutdown windows
'Useage: dll.osexitwindows("2") reboots windows

Dim bootans As Integer, bootvalue As Long
On Error Resume Next

Select Case BootMode
    Case 1 'Shutdown Windows
        bootans = vbYes
            If bootans = vbYes Then
                bootvalue = ExitWindowsEx(EWX_SHUTDOWN, 0&)
            End If
    Case 2 ' Reboot Windows
        bootans = MsgBox("Are you sure you want to reboot windows?", vbQuestion Or vbYesNo, "Reboot Windows")
            If bootans = vbYes Then
                bootvalue = ExitWindowsEx(EWX_REBOOT, 0&)
            End If
    Case Else
End Select
End Function

Function FreeProcess()
'Info: Unfreezes a locked loop or subroutine
'Useage: dll.freeprocess
Dim Process%
Do: DoEvents
    Process = Process + 1
    If Process = 50 Then Exit Do
Loop
End Function

Public Function FileDirCheck(ByVal sDirName As String) As Boolean
'Info: Returns true if path exists else it returns false
'Useage: dll.filedircheck("Filename") Will return true or false
Dim sDir As String
On Error Resume Next
    FileDirCheck = False
    sDir = Dir$(sDirName, vbDirectory)
        If (Len(sDir) > 0) And (Err = 0) Then
            FileDirCheck = True
        End If
End Function

Public Function FileSize(FilePathName As String) As String
'Info: Returns size of a file
'USAGE: Label1.Caption = FileSize("C:\Stuff.exe")
Dim Str As Single
    Str = FileLen(FilePathName)
    FileSize = File_ByteConversion(Str)
End Function

Public Function OsUptime()
'Info: Returns current uptime of your personal computer
'Useage: text1.text dll.osuptime
Dim Secs, Mins, Hours, Days
Dim TotalMins, TotalHours, TotalSecs, TempSecs
Dim CaptionText
    TotalSecs = Int(GetTickCount / 1000)
    Days = Int(((TotalSecs / 60) / 60) / 24)
    TempSecs = Int(Days * 86400)
    TotalSecs = TotalSecs - TempSecs
    TotalHours = Int((TotalSecs / 60) / 60)
    TempSecs = Int(TotalHours * 3600)
    TotalSecs = TotalSecs - TempSecs
    TotalMins = Int(TotalSecs / 60)
    TempSecs = Int(TotalMins * 60)
    TotalSecs = (TotalSecs - TempSecs)
        If TotalHours > 23 Then
            Hours = (TotalHours - 23)
        Else
            Hours = TotalHours
        End If
        
        If TotalMins > 59 Then
            Mins = (TotalMins - (Hours * 60))
        Else
            Mins = TotalMins
        End If

    CaptionText = Days & " Days, " & Hours & " Hours, " & Mins & " Minutes, " & TotalSecs & " seconds"
    OsUptime = CaptionText
End Function

Public Function Show()
'Info: Just displays the about dialog
'Useage: dll.show
    frmAbout.Show
End Function
