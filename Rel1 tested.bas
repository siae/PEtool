Attribute VB_Name = "Module1"
'***********************************************************
'Tool: Consolidate multiple DP submissions into a single tracker
'Release: 24/03/2017
'To Do: replicate for FTTB Tracker & Exception sheets
'Email: edmundsia@nbnco.com.au
'***********************************************************
Option Explicit
Public Const gMaxoldmail As Integer = 2
Public gSdate As Date
Public gEdate As Date
Public gk As Long

Public gRootfolder As String
Public tgtTracker As String
Public tgtTrackerPath As String
Public gSht1start As String
Public gSht2copy As String

Public rgEmaillog As Range

Public Const byKey As Integer = 1
Public Const byItem As Integer = 2

Public gHeader_rownum As Scripting.Dictionary
Public gFileDict As Scripting.Dictionary
Public gAttachmentDict As Scripting.Dictionary

Public Const DParr2 = "BROADSPECTRUM:DECON:DOWNER:FULTONHOGAN:LENDLEASE:QCCOMMUNICATIONS:QC COMMS:SAPN:SERVICESTREAM:SERVICE STREAM:VISIONSTREAM:VPL:WBHO"
Public Const Statearr2 = "QLD:ACT:VIC:NSW:WA:SA"
Public Const Regionarr2 = "NORTH:SOUTH"
Public Const Colarr2 = "X:O:U:P:Q:R"
'Dim DParr As Variant, Statearr As Variant, Regionarr As Variant, Colcolarr As Variant

Public DParr() As String
Public Statearr() As String
Public Regionarr() As String
Public Colarr() As String


Public Sub go_download()
Dim sTime As Double, s As String, i As Integer, k

Call initglobal

If MsgBox("Excel window may appear frozen." _
            & vbCrLf & "Continue ?", vbOKCancel, "Start Processing") = vbCancel Then
    Exit Sub
End If

sTime = Timer

If SavedEmailAtt("Build Tracker Power", "Inbox", _
            DateValue(gSdate), DateValue(gEdate), "", _
            DParr, Statearr, Regionarr) < 1 Then

    MsgBox "No files between " & gSdate & "-" & gEdate, vbInformation, "Finished!"
    Wait (1)
    Exit Sub
End If

Application.StatusBar = "Done: download email attachments"
Wait (1)

'ThisWorkbook.Worksheets("email-log").Select
'Application.ScreenUpdating = True

'If MsgBox("Continue to Populate Tracker ?" & _
        vbCrLf & "Elapsed (sec): " & Round(Timer - sTime, 1), vbOKCancel, "Next Step") = vbCancel Then
'    Exit Sub
'End If
'Application.ScreenUpdating = False

Call BoosterOn("Validate email attachments ...")

Call Keep_validfiles(gFileDict)

'********************************
go_checkhdr_populate

'********************************
Call BoosterOff

Windows(tgtTracker).Activate
s = Left(tgtTrackerPath, Len(tgtTrackerPath) - 5) & " " & Format(Now(), "ddmmyyyy") & ".xlsx"

On Error GoTo saveNewtracker:
If Len(Dir(s, vbNormal)) > 0 Then Kill s

saveNewtracker:
ActiveWorkbook.SaveAs FileName:=s, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWorkbook.Close

Application.StatusBar = "All done in " & Round(Timer - sTime, 1) & "s. " & s
MsgBox "Result file: " & s


graceful_exit:
Exit Sub


ErrHandler:
MsgBox "An unexpected error has occurred." _
     & vbCrLf & "Error Number: " & Err.Number _
     & vbCrLf & "Error Description: " & Err.Description _
     , vbCritical, "Error!"
Resume graceful_exit

End Sub



Public Function initglobal()
Dim sht As Object

ReDim Statearr(0 To 5) 'Regionarr(5) =>6 items
ReDim Regionarr(0 To 1) 'Regionarr(1) =>2 items
ReDim DParr(0 To 9)
ReDim Colarr(0 To 5)

Application.StatusBar = ""
Wait (0.5)

DParr = Split(DParr2, ":")
Statearr = Split(Statearr2, ":")
Regionarr = Split(Regionarr2, ":")
Colarr = Split(Colarr2, ":")

Set sht = ThisWorkbook.Worksheets("Start")
gEdate = DateAdd("d", 1, DateValue(sht.Range("E6")))
gSdate = DateValue(sht.Range("E7"))  'DateSerial(yyyy, mm, dd)

gSht1start = "A6"  'Tracker header start
gSht2copy = "A6"  'DP file header start

gRootfolder = Application.ThisWorkbook.Path & "\"

tgtTracker = sht.Range("E10")
tgtTrackerPath = gRootfolder & tgtTracker

Call Make_subfolder(Statearr)

Call PopulateDict

'For Each key In gHeader_rownum.Keys
'    Debug.Print key, gHeader_rownum(key)
'Next key

Set sht = Nothing
End Function



'adapted from http://www.rondebruin.nl/win/s1/outlook/saveatt.htm
Private Function SavedEmailAtt(Sharedmailbox As String, mailfld As String, _
                    startDate As Date, endDate As Date, ArchiveFld As String, _
                    ByVal DParr As Variant, ByVal Statearr As Variant, ByVal Regionarr As Variant) _
                    As Integer
                    
Dim Ns As Namespace, myFolder As MAPIFolder, subFolder As MAPIFolder
Dim myItems As Outlook.Items, oFolder As Outlook.Folder, oAtmt As Outlook.Attachment
Dim oItem
Dim i As Integer, k As Integer, f As String, s As String, NewFile As String
Dim mystate As String, mydp As String, myregion As String, NewFile2 As String
Dim sht As Worksheet   'rgLog As Range,


GoClearFolder (Statearr)

If mailfld = "" Then
    mailfld = "Inbox"
End If

Set Ns = GetNamespace("MAPI")
Set myFolder = Ns.GetDefaultFolder(olFolderInbox)

For Each oFolder In Ns.Folders
    'Debug.Print oFolder.Name
    If oFolder.Name = Sharedmailbox Then
        Debug.Print "Found " & oFolder.Name
        Set myFolder = oFolder.Folders.Item(mailfld)
    End If
Next

Set subFolder = myFolder
Set myItems = subFolder.Items

If subFolder.Items.Count = 0 Then
    MsgBox "No email found in: " & Sharedmailbox & "\" & mailfld, _
           vbInformation, "Nothing Found"
    Set subFolder = Nothing
    Set myFolder = Nothing
    Set Ns = Nothing
    Exit Function
End If

If ArchiveFld = "" Then ArchiveFld = gRootfolder & "z0 archive\"

Set sht = ThisWorkbook.Worksheets("email-log")
s = "A" & sht.Cells(sht.Rows.Count, "A").End(xlUp).Row

Set rgEmaillog = ThisWorkbook.Worksheets("email-log").Range(s)

myItems.Sort "[ReceivedTime]", True  'sort Desc=True


'init
Set gFileDict = New Scripting.Dictionary        'att and state-region-DP
Set gAttachmentDict = New Scripting.Dictionary  'att and email-log row #

i = 0
k = 0

For Each oItem In myItems
    'Debug.Print (i & "=" & oItem.SenderName & "=" & oItem.ReceivedTime)
    Application.StatusBar = oItem.ReceivedTime & ": " & oItem.Subject
    Wait (0.5)
    
    If (oItem.ReceivedTime < startDate) Then k = k + 1
    If k > gMaxoldmail Then GoTo skip_oldmail:
    
    If (oItem.ReceivedTime > startDate And oItem.ReceivedTime < endDate) Then  'inclusive start/end dates
        For Each oAtmt In oItem.Attachments
            f = UCase(oAtmt.FileName)
            
            Select Case Right(f, 4)
                Case ".XLS", "XLSX", "XLSM"
                    
                    i = i + 1

                    mydp = matchedregex(f, DParr)
                    
                    Select Case mydp
                        Case "VPL"
                            mydp = "VISIONSTREAM"
                        Case "QC COMMS"
                            mydp = "QCCOMMUNICATIONS"
                        Case "SERVICE STREAM"
                            mydp = "SERVICESTREAM"
                    End Select
                    
                    mystate = matchedregex(f, Statearr)
                    
                    myregion = matchedregex(f, Regionarr)
                    'Debug.Print mystate, myregion, mydp, f
                    
                    NewFile = ""
                    NewFile2 = ""
                    
                    If mystate = "na" Or mydp = "na" Then
                        NewFile = gRootfolder & "Exception\" & GetPartOfFilePath(f, "fNameExt")
                    
                    Else
                        NewFile2 = mystate & "-" & myregion & "-" & mydp & "-" & _
                                    Format(oItem.ReceivedTime, "yyyymmdd-hhmm") & "." & GetPartOfFilePath(f, "fExt")
                                    
                        gFileDict.Add key:=NewFile2, Item:=mystate & "-" & myregion & "-" & mydp
                        
                        gAttachmentDict.Add key:=NewFile2, Item:=i
                        
                        NewFile = gRootfolder & mystate & "\" & NewFile2
                    
                    End If
                    
                    oAtmt.SaveAsFile NewFile
                                       
                    rgEmaillog.Offset(i, 0) = NewFile2  'attachment saved name
                    'rgEmaillog.Offset(i, 1) = attachment status
                    rgEmaillog.Offset(i, 2) = mystate
                    rgEmaillog.Offset(i, 3) = myregion
                    rgEmaillog.Offset(i, 4) = mydp
                    rgEmaillog.Offset(i, 5) = oItem.ReceivedTime
                    rgEmaillog.Offset(i, 6) = oItem.Sender
                    rgEmaillog.Offset(i, 7) = oItem.SenderEmailAddress
                    rgEmaillog.Offset(i, 8) = oAtmt.FileName
                    rgEmaillog.Offset(i, 9) = oItem.Subject
                                        
            End Select
        Next oAtmt
    End If
Next oItem


skip_oldmail:
SavedEmailAtt = i
Application.StatusBar = "Saved " & i & " files"


graceful_exit:
Set sht = Nothing
Set oAtmt = Nothing
Set oItem = Nothing
Set oFolder = Nothing
Set myItems = Nothing
Set subFolder = Nothing
Set myFolder = Nothing
Set Ns = Nothing
Exit Function


ErrHandler:
MsgBox "An unexpected error has occurred." _
     & vbCrLf & "Error Number: " & Err.Number _
     & vbCrLf & "Error Description: " & Err.Description _
     , vbCritical, "Error!"
Resume graceful_exit

End Function



Private Function GoClearFolder(Fldarr As Variant)
Dim n As Long
For n = LBound(Fldarr) To UBound(Fldarr)
    On Error Resume Next
    Kill gRootfolder & Fldarr(n) & "\*.*"
    On Error GoTo 0
Next n

On Error Resume Next
Kill gRootfolder & "Exception\*.*"
On Error GoTo 0

End Function



Private Function matchedregex(mystr As String, regexarr As Variant) As String
Dim s As String, n As Long, matches As MatchCollection, m As Object
Dim RegX As VBScript_RegExp_55.RegExp   'early binding
'Dim RegX As Object                     'late binding

Set RegX = New VBScript_RegExp_55.RegExp    'early binding
'Set RegX = CreateObject("VBScript.RegExp") 'late binding

matchedregex = "na"

For n = LBound(regexarr) To UBound(regexarr)
    s = regexarr(n)
    With RegX
      .Pattern = regexarr(n)
      .Global = True
    End With
     
    Set matches = RegX.Execute(mystr)
       
    For Each m In matches
        matchedregex = m.Value
        Exit Function
    Next m
Next n

End Function



Public Function go_checkhdr_populate()
Dim srcFile As String, bool As Boolean, n As Long, s As String, f As String, s1 As String
Dim wb1 As Workbook, sht1 As Worksheet, rg1 As Range, wb2 As Workbook
Dim newColRow As String, result As String

Call initglobal

'Call BoosterOn("Checking DP file headers ...")
'sTime = Timer

On Error GoTo nextStep:
Windows(tgtTracker).Activate
ActiveWorkbook.Close savechanges:=False

nextStep:
'On Error GoTo ErrHandler:
Set wb1 = Workbooks.Open(FileName:=tgtTrackerPath, UpdateLinks:=0)


For n = LBound(Statearr) To UBound(Statearr)
    s = gRootfolder & Statearr(n) & "\"  '"C:\00 Projects\BPE Fttx\Power Energisation\QLD"
    f = Dir(gRootfolder & Statearr(n) & "\*.xls*", vbNormal)
    
    If Len(f) = 0 Then
        Application.StatusBar = "WARNING: No files for " & Statearr(n)
        Wait (1)
    End If
    
    Set sht1 = wb1.Worksheets(Statearr(n) & " FTTN Tracking Register")  'WA FTTN Tracking Register"
    Set rg1 = sht1.Range(gSht1start) 'A6
    sht1.Select
        
    newColRow = gSht1start
    
    'wb1 = Tracker   wb2 = DP file
    Do While Len(f) > 0
        s1 = s & f
        Set wb2 = Workbooks.Open(FileName:=s1, UpdateLinks:=0)
        
        bool = hdr_ok(wb2)   'beware wb2=nothing, file was closed
        
        If Not bool Then  'wb2=nothing, wb2 is closed if DP header is shorter
            Application.StatusBar = "** Header check FAILED. Closing " & f
            Wait (1)
            
            'cannot copy of move files, f=Dir will fail
            'Call Copyfile2Folder(s1, gRootfolder & "Exception\")
        
        Else
            'newColRow = sht1ColRow, is updated & used to paste next DP file data
            result = GoCopyPaste(wb2, gSht2copy, Colarr(n), sht1, newColRow)
            Application.StatusBar = result & ": " & f
            Wait (0.5)
        End If
        
        f = Dir
    Loop
Next


graceful_exit:
Application.StatusBar = "Done: Check header & Populate"
Wait (1)


Set wb1 = Nothing
Set sht1 = Nothing
Set rg1 = Nothing
Exit Function


ErrHandler:
MsgBox "An unexpected error has occurred." _
     & vbCrLf & "Error Number: " & Err.Number _
     & vbCrLf & "Error Description: " & Err.Description _
     , vbCritical, "Error!"
Resume graceful_exit:

End Function



'copy from wb2 DP file to wb1 Tracker
Private Function GoCopyPaste(wb2 As Workbook, sht2ColRow As String, ByVal eCol As String, _
                                sht1 As Object, sht1ColRow As String) As String

Dim sht2 As Worksheet, rg2 As Range   'wb2 As Workbook,
Dim sht1Col As String, sht2Col As String, sht1Row As Long, sht2Row As Long
Dim i As Long, s As String

GoCopyPaste = "FAIL"

Application.StatusBar = "Populating " & wb2.Name
Wait (0.5)

'copy from wb2 DP file to wb1 Tracker
'On Error GoTo ErrHandler:

'Set wb2 = Workbooks.Open(FileName:=srcFile, UpdateLinks:=0)
Set sht2 = wb2.Worksheets("FTTN Tracking Register")

'must unhide column & remove filter b4 copy data
sht2.Rows("4:4").EntireColumn.Hidden = False

On Error Resume Next
sht2.ShowAllData   'ActiveSheet.ShowAllData
On Error GoTo 0


sht2Col = Left(sht2ColRow, 1)  'A
sht2Row = Int(Mid(sht2ColRow, 2, Len(sht2ColRow)))  '6

Set rg2 = sht2.Range(sht2ColRow)
i = 1
s = Trim(rg2)
Do While Len(s) > 0
    s = Trim(rg2.Offset(i, 0))
    i = i + 1
Loop
'lastrow = i + sht2Row - 2
s = sht2ColRow & ":" & eCol & (i + sht2Row - 2)

Set rg2 = sht2.Range(s)
rg2.Copy

'rg2.Copy Destination:=sht1.Range(sColRow)
'sht1.rg1.Resize(rg2.Rows.Count, rg2.Columns.Count).Value = rg2.Value

sht1.Range(sht1ColRow).PasteSpecial Paste:=xlPasteValues


Application.DisplayAlerts = False

wb2.Close savechanges:=False
Application.DisplayAlerts = True


sht1Col = Left(sht1ColRow, 1)
sht1Row = Int(Mid(sht1ColRow, 2, Len(sht1ColRow)))

If sht1Row > 6 Then
    sht1ColRow = sht1Col & sht1Row + i - 1
    'sht1ColRow = sht1Col & (i + sht2Row - 1)
Else
    sht1ColRow = sht1Col & (i + sht2Row - 2) + 1  'updated. New value is the row to paste new DP file data
End If


graceful_exit:
'Set wb2 = Nothing
Set sht2 = Nothing
Set rg2 = Nothing

GoCopyPaste = "SUCCESS"
Exit Function

ErrHandler:
MsgBox "An unexpected error has occurred." _
     & vbCrLf & "Error Number: " & Err.Number _
     & vbCrLf & "Error Description: " & Err.Description _
     , vbCritical, "Error!"
Resume graceful_exit:

End Function



Private Function hdr_ok(wb2 As Workbook) As Boolean
Dim sht2 As Worksheet, sht3 As Worksheet, sht4 As Worksheet  'wb2 As Workbook,
Dim rg2 As Range, rg3 As Range, rg4 As Range, lastcol As Long, colX As Long, colY As Long
Dim s As String, s1() As String, key As String, v As String, i As Long, X As String, Y As String
Dim bHdr_deleted As Boolean

bHdr_deleted = False
hdr_ok = False

Set sht3 = ThisWorkbook.Worksheets("headers")
Set sht4 = ThisWorkbook.Worksheets("header-log")

s = "A" & sht4.Cells(sht4.Rows.Count, "A").End(xlUp).Row
Set rg4 = sht4.Range(s)


Set rg3 = sht3.Range("B1")

Application.StatusBar = "Check header: " & wb2.Name
Wait (0.5)

'Set wb2 = Workbooks.Open(FileName:=srcFile, UpdateLinks:=0)
Set sht2 = wb2.Worksheets("FTTN Tracking Register")
Set rg2 = sht2.Range("A4")
sht2.Select

'dont need to unhide column, only unhide to be able to see issues
sht2.Rows("4:4").EntireColumn.Hidden = False


Cells.Select
Application.CutCopyMode = False
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
rg2.Select

s = wb2.Name
s1 = Split(s, "-")  'state info
key = Trim(s1(0)) & " FTTN Tracking Register"

If gHeader_rownum.Exists(key) Then
    v = gHeader_rownum(key)  'Debug.Print key, gHeader_rownum(key)
Else
    MsgBox ("Header check failed: " & key & " not found")
    GoTo failexit:  'wb2=nothing, file is closed
End If

'sht3 Tracker header reference sht, sht2 DPfile header
'hidden column can be checked, but data not copied
'sht2.Rows("4:4").EntireColumn.Hidden = False

i = 0
gk = 0 'global

'sht3 Tracker header
X = rg3.Offset(v, i)  'offset v from Cells(1,0) = row v+1
If Len(X) > 30 Then X = Left(X, 30)
colX = sht3.Cells(v + 1, sht3.Columns.Count).End(xlToLeft).Column - 1

'sht2 DP file header
Y = rg2.Offset(0, i)
If Len(Y) > 30 Then Y = Left(X, 30)
colY = sht2.Cells(4, sht2.Columns.Count).End(xlToLeft).Column

Do While Len(X) > 0
    If colX > colY Then
        gk = gk + 1
        'MsgBox "Critical Error: DP has less header columns", vbExclamation
        rg4.Offset(gk, 0) = s
        rg4.Offset(gk, 1) = "Expect " & colX & " columns"
        rg4.Offset(gk, 2) = "Detected " & colY
        rg4.Offset(gk, 3) = "DP missing 1 or more header column"
        
        GoTo failexit:  'wb2=nothing, file is closed
    End If

    If StrComp(X, Y, vbTextCompare) <> 0 Then 'case insensitive
        gk = gk + 1
        rg4.Offset(gk, 0) = s
        rg4.Offset(gk, 1) = X
        rg4.Offset(gk, 2) = Y
        'Debug.Print X, rg3.Offset(v, i).Address, Y, rg2.Offset(0, i).Address
        
        Application.StatusBar = s & ">>  Deleting " & Y
        Wait (0.5)

        Call delete_badcolumn(sht2, col_letter(i + 1))
        
        rg4.Offset(gk, 3) = "Deleted"
        bHdr_deleted = True
        'wb2.Save
        
        i = 0
    Else
        i = i + 1
    End If
    
    X = rg3.Offset(v, i)
    If Len(X) > 30 Then X = Left(X, 30)
    colX = sht3.Cells(v + 1, sht3.Columns.Count).End(xlToLeft).Column - 1 'col A is state

    Y = rg2.Offset(0, i)
    If Len(Y) > 30 Then Y = Left(X, 30)
    colY = sht2.Cells(4, sht2.Columns.Count).End(xlToLeft).Column
    
Loop

hdr_ok = True

If bHdr_deleted Then wb2.Save


graceful_exit:  'wb2 is not closed
Set sht3 = Nothing
Set sht4 = Nothing
Set rg2 = Nothing
Set rg3 = Nothing
Set rg4 = Nothing
Exit Function


failexit:  'DP has less column, wb2=nothing
Application.DisplayAlerts = False

wb2.Close savechanges:=False
Application.DisplayAlerts = True
GoTo graceful_exit:

End Function



Private Function delete_badcolumn(sht2 As Object, col As String)
Dim s As String, rg As Object

s = col & "4"
Set rg = sht2.Range(s)
rg.EntireColumn.Delete

Set rg = Nothing
End Function


Private Function col_letter(colnum As Long) As String
Dim vArr
vArr = Split(Cells(1, colnum).Address(True, False), "$")
col_letter = vArr(0)
End Function

'Public Sub go_populate_tracker(Statearr As Variant, Colarr As Variant, tgtTrackerPath As String)
'Dim DParr As Variant, Statearr As Variant, Regionarr As Variant, Colcolarr As Variant
'DParr = Array("BROADSPECTRUM", "DECON", "DOWNER", "FULTONHOGAN", "LENDLEASE", _
                    "QCCOMMUNICATIONS", "SAPN", "SERVICESTREAM", "VISIONSTREAM", "WBHO")
'Statearr = Array("QLD", "ACT", "VIC", "NSW", "WA", "SA")
'Regionarr = Array("NORTH", "SOUTH")
'Colarr = Array("X", "O", "U", "P", "Q", "R")



Public Function GetPartOfFilePath(sfullpath As String, sOpt As String) As String
GetPartOfFilePath = ""

Select Case sOpt
    Case "fFullpath"
        'ActiveWorkbook's File Path
        GetPartOfFilePath = ActiveWorkbook.FullName
    
    Case "fPathNoExt"
        'Take Off File Extension
        GetPartOfFilePath = Left(sfullpath, InStrRev(sfullpath, ".") - 1)
        
    Case "fExt"
        'get File Ext
        GetPartOfFilePath = Right(sfullpath, Len(sfullpath) - InStrRev(sfullpath, "."))
        
    Case "fNameExt"
        'get File Name & Ext
        GetPartOfFilePath = Right(sfullpath, Len(sfullpath) - InStrRev(sfullpath, "\"))
        
    Case "fName"
        'get File Name without Ext
        GetPartOfFilePath = Mid(sfullpath, InStrRev(sfullpath, "\") + 1, InStrRev(sfullpath, ".") - InStrRev(sfullpath, "\") - 1)
End Select

End Function



Public Function Wait(lSec As Single)
Dim time1 As Variant
'Timer = Single representing the number of seconds elapsed since midnight
time1 = Timer + lSec
    
Do Until Timer > time1
    DoEvents
Loop
End Function



Private Function BoosterOn(msg As String)
'oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = msg

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
'Application.DisplayStatusBar = False
Application.EnableEvents = False
End Function



Private Function BoosterOff()
Application.EnableEvents = True
Application.DisplayStatusBar = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

'Application.StatusBar = ""
'Application.DisplayStatusBar = oldStatusBar
End Function



Private Function Make_subfolder(ByVal Statearr As Variant)
Dim n As Long
Dim fso As Scripting.FileSystemObject   'early binding
'Dim fso As Object                      'late binding

Set fso = New Scripting.FileSystemObject
'Set fso = CreateObject("Scripting.FileSystemObject")
    
For n = LBound(Statearr) To UBound(Statearr)
    If Not fso.FolderExists(gRootfolder & Statearr(n)) Then
        fso.CreateFolder gRootfolder & Statearr(n)
    End If
Next

If Not fso.FolderExists(gRootfolder & "Exception\") Then
    fso.CreateFolder gRootfolder & "Exception\"
End If

Set fso = Nothing
End Function



Private Function PopulateDict()
Dim i As Integer, k As String, sht3 As Object

Set gHeader_rownum = New Scripting.Dictionary  'headers sht: state shtname vs rownum

Set sht3 = ThisWorkbook.Worksheets("headers")

i = 0
With sht3
    k = .Range("A1").Offset(i, 0)
    
    Do While Len(k) > 0
        gHeader_rownum.Add key:=k, Item:=i
        
        i = i + 1
        k = .Range("A1").Offset(i, 0)
    Loop
End With

Set sht3 = Nothing
End Function



Private Function Keep_validfiles(gFileDict As Object)
Dim s As String, rgLog As Range, sht As Worksheet
Dim sKey, sItem, ref, v

Call SortDict(gFileDict, byKey, "DESC")

ref = ""

For Each sKey In gFileDict.Keys
    'Debug.Print sKey, gFileDict(sKey), ref
    sItem = gFileDict(sKey)
    
    If sItem = ref Then
        s = gRootfolder & Split(sKey, "-")(0) & "\" & sKey
        
        If FileExist(s) Then
            If gAttachmentDict.Exists(sKey) Then
                v = gAttachmentDict(sKey)
                'Debug.Print "**" & sKey, gFileDict(sKey), ref
            Else
                v = 0
            End If
            rgEmaillog.Offset(v, 1) = "Removed"

            SetAttr s, vbNormal  'in case read-only file
            Kill (s)
            
        End If
    Else
        ref = sItem
    End If
    
Next sKey

Application.StatusBar = "Done: kept valid attachments"
Wait (1)
End Function


Private Function FileExist(ByVal myFile As String) As Boolean
    If Len(myFile) = 0 Then
        FileExist = False
        Exit Function
    End If
    FileExist = (Dir(myFile) <> "")
End Function



Private Function SortDict(oDict As Object, iSort As Integer, sOrder As String) 'sort Desc
Dim sDict() As String, oKey, sKey, sItem
Dim X As Integer, Y As Integer, Z As Integer

Z = oDict.Count

'if dict count < 1
If Z > 1 Then
    ReDim sDict(Z, 2)
    X = 0
    
    For Each oKey In oDict
        sDict(X, byKey) = CStr(oKey)
        sDict(X, byItem) = CStr(oDict(oKey))
        X = X + 1
    Next
    
    For X = 0 To (Z - 2)
        For Y = X To (Z - 1)
      
            Select Case sOrder
                Case "DESC":
                    If StrComp(sDict(X, iSort), sDict(Y, iSort), vbTextCompare) < 0 Then  'sort desc
                        sKey = sDict(X, byKey)
                        sItem = sDict(X, byItem)
                        sDict(X, byKey) = sDict(Y, byKey)
                        sDict(X, byItem) = sDict(Y, byItem)
                        sDict(Y, byKey) = sKey
                        sDict(Y, byItem) = sItem
                    End If
            
                Case "ASC":
                    If StrComp(sDict(X, iSort), sDict(Y, iSort), vbTextCompare) > 0 Then  'sort ASC
                        sKey = sDict(X, byKey)
                        sItem = sDict(X, byItem)
                        sDict(X, byKey) = sDict(Y, byKey)
                        sDict(X, byItem) = sDict(Y, byItem)
                        sDict(Y, byKey) = sKey
                        sDict(Y, byItem) = sItem
                    End If
            End Select
        Next
    Next
    
    oDict.RemoveAll
    
    ' repopulate dictionary with sorted information
    For X = 0 To (Z - 1)
      oDict.Add sDict(X, byKey), sDict(X, byItem)
    Next
End If

End Function



Private Function Copyfile2Folder(ByVal srcFile As String, ByVal sFld As String)
Dim s As String
Dim fso As Scripting.FileSystemObject   'early binding
'Dim fso As Object                      'late binding

Set fso = New Scripting.FileSystemObject                 'early binding
'Set fso = CreateObject("scripting.filesystemobject")    'late binding

s = GetPartOfFilePath(srcFile, "fNameExt")

If sFld = "" Then sFld = gRootfolder & "Exception\"

If Right(sFld, 1) <> "\" Then
    sFld = sFld & "\"
End If

If fso.FolderExists(sFld) = False Then
    fso.CreateFolder (sFld)
End If
    
On Error GoTo ErrHandler:
If Len(Dir(srcFile, vbNormal)) = 0 Then GoTo ErrHandler:  'doesnt exist

If Len(Dir(sFld & s, vbNormal)) > 0 Then Kill s   'remove existing copy in new folder

fso.CopyFile Source:=srcFile, Destination:=sFld   'copy to new folder
'fso.MoveFile Source:=srcFile, Destination:=sFld   'move to new folder

Application.StatusBar = s & " copied to Exception folder"
Wait (1)

Exit Function

ErrHandler:
MsgBox "File move failed. " & srcFile & " was not found.", vbExclamation

End Function



Sub test()
Call initglobal
'GoClearFolder (Statearr)
Call go_checkhdr_populate
End Sub


'initglobal
'GoClearFolder
'go_download
'SavedEmailAtt
'go_check_headers
'hdr_ok
'GoCopyPaste
'PopulateDict
'SortDict
'Keep_validfiles
'col_letter
'delete_badcolumn
'Copyfile2Folder

