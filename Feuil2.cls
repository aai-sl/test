VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub saveThisWorkbook()

Dim depotPath As String
Dim resultsPath As String
Dim timestamp As String

timestamp = generateTimeStamp()
depotPath = findVariable("Chemin de depot de fichier pour génération")
resultsPath = findVariable("Archive de resultats")

On Error Resume Next
MkDir resultsPath & "\" & timestamp
On Error GoTo 0

resultsPath = resultsPath & "\" & timestamp


Application.DisplayAlerts = False
ActiveWorkbook.SaveCopyAs Filename:="" & depotPath & "\" & ActiveWorkbook.Name
Application.DisplayAlerts = True

retVal = Shell("explorer.exe " & resultsPath, vbNormalFocus)


End Sub

Function findVariable(findMe As String)
    Dim match As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    Set match = ws.Cells.Find(findMe)
    findVariable = match.Offset(, 1).Value
  
End Function

Function generateTimeStamp()
    
Dim DT
DT = Format(Now, "yyyymmdd_hhmmss")
    
    
    Dim match As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    Set match = ws.Cells.Find("Timestamp Actuel")
    match.Offset(, 1).Value = DT
    generateTimeStamp = DT
End Function

