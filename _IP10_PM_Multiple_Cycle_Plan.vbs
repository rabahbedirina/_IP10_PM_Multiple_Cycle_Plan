' Script Name: _IP10_PM_Multiple_Cycle_Plan.vbs
' Description: Automated SAP PM IP10 Multiple Cycle Plan Execution
' Version:     1.0.0
' Author:      Rabah BEDIRINA
' GitHub:      https://github.com/rabahbedirina/
' ==============================================================================

Option Explicit

Dim shell, fso, file
Dim csvPath, csvLine, fields
Dim MaintenancePlan, StartDate, Sequence
Dim TAB_COUNT
Dim i

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

csvPath = shell.CurrentDirectory & "\DATA__IP10_Multiple_Cycle_Plan.csv"

If Not fso.FileExists(csvPath) Then
    csvPath = shell.CurrentDirectory & "\DATA_IP10_Multiple_Cycle_Plan.csv"
End If

If Not fso.FileExists(csvPath) Then
    MsgBox "CSV file not found:" & vbCrLf & shell.CurrentDirectory & "\DATA__IP10_Multiple_Cycle_Plan.csv" & vbCrLf & _
           shell.CurrentDirectory & "\DATA_IP10_Multiple_Cycle_Plan.csv", vbCritical, "Error"
    WScript.Quit
End If

Set file = fso.OpenTextFile(csvPath, 1)

' Skip CSV header
If Not file.AtEndOfStream Then
    file.ReadLine
End If

MsgBox "Open SAP and make sure you re ready." & vbCrLf & _
       "Click OK to start.", _
       vbInformation, "SAP Automation"
Do Until file.AtEndOfStream

    csvLine = file.ReadLine

    If Trim(csvLine) <> "" Then

        fields = Split(csvLine, ";")

        If UBound(fields) >= 1 Then

            MaintenancePlan = Trim(fields(0))
            StartDate = Trim(fields(1))
            Sequence = CLng(Trim(fields(2)))


            shell.AppActivate "SAP"
            WScript.Sleep 500

            shell.SendKeys "^/"
            WScript.Sleep 500

            shell.SendKeys "/nIP10"
            shell.SendKeys "{ENTER}"

            WScript.Sleep 500

            shell.SendKeys MaintenancePlan
            shell.SendKeys "{ENTER}"
            
            WScript.Sleep 500
            
            shell.SendKeys "^{F7}"
            
            if Sequence = 1 Then                
            WScript.Sleep 100
            shell.SendKeys "+{TAB}"
            WScript.Sleep 100
            else
            TAB_COUNT = Sequence - 1
            For i = 1 To TAB_COUNT
                WScript.Sleep 100
                shell.SendKeys "{TAB}"
                WScript.Sleep 100
            Next
            End If

            shell.SendKeys " "
            WScript.Sleep 250
            
            shell.SendKeys "{TAB}"
            WScript.Sleep 250

            shell.SendKeys StartDate
            shell.SendKeys "{TAB}"
            WScript.Sleep 250

            shell.SendKeys "07"           
            WScript.Sleep 500
            
            shell.SendKeys "{ENTER}"
            WScript.Sleep 500         
            shell.SendKeys "{ENTER}"
            WScript.Sleep 500
            shell.SendKeys "^s"

            WScript.Sleep 500

        End If

    End If

Loop

file.Close

MsgBox "All rows have been processed.", _
       vbInformation, "SAP Automation"

Set file = Nothing
Set fso = Nothing
Set shell = Nothing
