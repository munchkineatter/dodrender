' ===========================================
' ExportGameState Module for Render Hosting
' Add this to your VBA project (Insert > Module)
' ===========================================

Option Explicit

' *** IMPORTANT: Update this URL after deploying to Render! ***
' It will look like: https://your-app-name.onrender.com
Public Const RENDER_URL As String = "https://YOUR-APP-NAME.onrender.com"

' Send eliminated prizes to Render server
Public Sub ExportGameStateToJSON()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gameboard") ' Change if your sheet has a different name
    
    Dim eliminated As String
    Dim i As Integer
    Dim cellValue As Variant
    Dim isEliminated As Boolean
    
    eliminated = ""
    
    ' Check each prize in column M (13) to see if it's strikethrough
    For i = 1 To 17
        cellValue = ws.Cells(9 + i, 13).Value
        isEliminated = ws.Cells(9 + i, 13).Font.Strikethrough
        
        If isEliminated Then
            If Len(eliminated) > 0 Then
                eliminated = eliminated & ", "
            End If
            eliminated = eliminated & CStr(cellValue)
        End If
    Next i
    
    ' Build JSON
    Dim json As String
    json = "{""eliminatedPrizes"": [" & eliminated & "]}"
    
    ' Send to Render server
    SendToRender "/api/update", json
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error exporting game state: " & Err.Description
End Sub

' Reset the display (call when starting a new game)
Public Sub ResetGameDisplay()
    On Error GoTo ErrorHandler
    
    SendToRender "/api/reset", "{}"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error resetting game display: " & Err.Description
End Sub

' Send HTTP POST request to Render server
Private Sub SendToRender(endpoint As String, jsonData As String)
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    url = RENDER_URL & endpoint
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonData
    
    If http.Status = 200 Then
        Debug.Print "Successfully sent to Render: " & endpoint
    Else
        Debug.Print "Error from Render: " & http.Status & " - " & http.statusText
    End If
    
    Set http = Nothing
    Exit Sub
    
ErrorHandler:
    Debug.Print "HTTP Error: " & Err.Description
    Set http = Nothing
End Sub

' Test function - eliminates first 3 prizes for testing
Public Sub TestEliminatePrizes()
    On Error GoTo ErrorHandler
    
    Dim json As String
    json = "{""eliminatedPrizes"": [1500, 2000, 2200]}"
    
    SendToRender "/api/update", json
    
    MsgBox "Test sent to Render: Eliminated $1,500, $2,000, $2,200", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure:" & vbCrLf & _
           "1. Render URL is correct in the code" & vbCrLf & _
           "2. Your internet connection is working" & vbCrLf & _
           "3. The Render service is running", vbExclamation
End Sub

' Wake up the Render server (free tier sleeps after inactivity)
Public Sub WakeUpServer()
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    url = RENDER_URL & "/health"
    
    http.Open "GET", url, False
    http.send
    
    If http.Status = 200 Then
        MsgBox "Server is awake and ready!", vbInformation
    Else
        MsgBox "Server may be starting up. Wait 30 seconds and try again.", vbExclamation
    End If
    
    Set http = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Could not reach server. It may be starting up (takes ~30 seconds on free tier)." & vbCrLf & _
           "Error: " & Err.Description, vbExclamation
    Set http = Nothing
End Sub

