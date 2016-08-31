' Usage:
'  CScript decompose.vbs <file> <path> [<stub>] [--no-stub]
'
' Converts all modules, classes, forms and macros from an all Access Project file (.adp) and Access Databases (.mdb)
' from <path>
' text and saves the results in separate files to decomposeDir.  Requires Microsoft Access.


Option Explicit

const acForm = 2
const acModule = 5
const acMacro = 4
const acReport = 3

const ForReading = 1, ForWriting = 2, ForAppending = 8

If (WScript.Arguments.Count < 2) Then
    WScript.Echo "CScript decompose.vbs <file> <path> [<stub>] [--no-stub]"
    WScript.Quit
End If

Dim projectFile, decomposeDir, stubFile, useStub
projectFile  = WScript.Arguments(0)
decomposeDir = WScript.Arguments(1)
stubFile = ""
useStub = True

If (WScript.Arguments.Count > 2) Then
    If (WScript.Arguments(2) = "--no-stub") Then
        useStub = False
    Else
        stubFile = WScript.Arguments(2)
        If (WScript.Arguments.Count > 3) Then
            If (WScript.Arguments(3) = "--no-stub") Then
                useStub = False
            End If
        End If
    End If
End If

' BEGIN CODE

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

exportModulesTxt projectFile, decomposeDir, stubFile, useStub

If (Err <> 0) and (Err.Description <> NULL) Then
    MsgBox Err.Description, vbExclamation, "Error"
    Err.Clear
End If

Function exportModulesTxt(projectFile, sExportpath, stubFile, useStub)

    Dim sADPFilename
    sADPFilename = fso.GetAbsolutePathName(projectFile)

    ' Build file and pathnames
    Dim myType, myName, myPath
    myType = fso.GetExtensionName(sADPFilename)
    myName = fso.GetBaseName(sADPFilename)
    myPath = fso.GetParentFolderName(sADPFilename)

    If (stubFile = "") Then
        stubFile = sExportpath & "\" & myName & "_stub." & myType
    End If

    Dim sStubADPFilename
    sStubADPFilename = fso.GetAbsolutePathName(stubFile)

    Dim myComponent
    Dim sModuleType
    Dim sTempname
    Dim sOutstring

    sExportpath = fso.GetAbsolutePathName(sExportpath)

    If (useStub) Then
        WScript.Echo "copy " & sADPFilename & " to " & sStubADPFilename & "..."
        On Error Resume Next
            fso.CreateFolder(sExportpath)
        On Error Goto 0
        fso.CopyFile sADPFilename, sStubADPFilename
    End If

    WScript.Echo "projectFile: " & sADPFilename
    WScript.Echo "sExportpath: " & fso.GetAbsolutePathName(sExportpath)
    WScript.Echo "stubFile   : " & sStubADPFilename
    WScript.Echo "useStub    : " & useStub
    WScript.Echo ""

    WScript.Echo "hold Shift and press Enter"
    WScript.StdIn.Read(2)

    WScript.Echo "starting Access..."
    Dim oApplication
    Set oApplication = CreateObject("Access.Application")

    Dim useADPFilename

    If (useStub) Then
        useADPFilename = sStubADPFilename
    Else
        useADPFilename = sADPFilename
    End If

    WScript.Echo "opening " & useADPFilename & " ..."
    If (Right(sADPFilename,4) = ".adp") Then
        oApplication.OpenAccessProject useADPFilename
    Else
        oApplication.OpenCurrentDatabase useADPFilename
    End If

    oApplication.Visible = false

    Dim dctDelete
    Set dctDelete = CreateObject("Scripting.Dictionary")
    WScript.Echo "exporting..."

    Dim myObj
    For Each myObj In oApplication.CurrentProject.AllForms
        WScript.Echo "  " & myObj.fullname & ".form"
        oApplication.SaveAsText acForm, myObj.fullname, sExportpath & "\" & myObj.fullname & ".form"
	Sanitize(sExportpath & "\" & myObj.fullname & ".form")
        oApplication.DoCmd.Close acForm, myObj.fullname
        dctDelete.Add "FO" & myObj.fullname, acForm
    Next
    For Each myObj In oApplication.CurrentProject.AllModules
        WScript.Echo "  " & myObj.fullname & ".bas"
        oApplication.SaveAsText acModule, myObj.fullname, sExportpath & "\" & myObj.fullname & ".bas"
        dctDelete.Add "MO" & myObj.fullname, acModule
    Next
    For Each myObj In oApplication.CurrentProject.AllMacros
        WScript.Echo "  " & myObj.fullname & ".mac"
        oApplication.SaveAsText acMacro, myObj.fullname, sExportpath & "\" & myObj.fullname & ".mac"
        dctDelete.Add "MA" & myObj.fullname, acMacro
    Next
    For Each myObj In oApplication.CurrentProject.AllReports
        WScript.Echo "  " & myObj.fullname & ".report"
        oApplication.SaveAsText acReport, myObj.fullname, sExportpath & "\" & myObj.fullname & ".report"
	Sanitize(sExportpath & "\" & myObj.fullname & ".report")
        dctDelete.Add "RE" & myObj.fullname, acReport
    Next

    If (useStub) Then
        WScript.Echo "deleting..."
        dim sObjectname
        For Each sObjectname In dctDelete
            WScript.Echo "  " & Mid(sObjectname, 3)
            oApplication.DoCmd.DeleteObject dctDelete(sObjectname), Mid(sObjectname, 3)
        Next
    End If

    oApplication.CloseCurrentDatabase
    If (useStub) Then
        oApplication.CompactRepair sStubADPFilename, sStubADPFilename & "_"
        fso.CopyFile sStubADPFilename & "_", sStubADPFilename
        fso.DeleteFile sStubADPFilename & "_"
    End If

    oApplication.Quit

End Function

Public Function getErr()
    Dim strError
    strError = vbCrLf & "----------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
               "From " & Err.source & ":" & vbCrLf & _
               "    Description: " & Err.Description & vbCrLf & _
               "    Code: " & Err.Number & vbCrLf
    getErr = strError
End Function

Public Sub Sanitize(fileName)
    Dim AggressiveSanitize 

    AggressiveSanitize = False
	
    Dim InFile, OutFile
    Set InFile = fso.OpenTextFile(fileName, ForReading)
    Set OutFile = fso.CreateTextFile(fileName & ".sanitize", True)	

    Dim txt
    Do Until InFile.AtEndOfStream
        txt = InFile.ReadLine
        If Left(txt, 10) = "Checksum =" Then
            ' Skip lines starting with Checksum
        ElseIf InStr(txt, "NoSaveCTIWhenDisabled =1") Then
            ' Skip lines containning NoSaveCTIWhenDisabled
        ElseIf InStr(txt, "Begin") > 0 Then
            If _
                InStr(txt, "PrtDevNames =") > 0 Or _
                InStr(txt, "PrtDevNamesW =") > 0 Or _
                InStr(txt, "PrtDevModeW =") > 0 Or _
                InStr(txt, "PrtDevMode =") > 0 _
                Then

                ' skip this block of code
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            ElseIf AggressiveSanitize And ( _
                InStr(txt, "dbLongBinary ""DOL"" =") > 0 Or _
                InStr(txt, "NameMap") > 0 Or _
                InStr(txt, "GUID") > 0 _
                ) Then

                ' skip this block of code
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            Else                       ' This line needs to be added
                OutFile.WriteLine txt  ' This line needs to be added
            End If                     ' This line needs to be added
        Else
            OutFile.WriteLine txt
        End If
    Loop
    OutFile.Close
    InFile.Close

    fso.CopyFile fileName & ".sanitize", fileName
    fso.DeleteFile fileName & ".sanitize"

End Sub
