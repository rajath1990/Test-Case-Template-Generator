Attribute VB_Name = "UserForm1Code"
Option Explicit


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Unload Me
Dim WB As Workbook

Sheet1.Activate
            range("A1").Select
            ActiveCell.End(xlDown).Select
            ActiveCell.Offset(1, 0).Select
            ActiveCell.Value = ActiveCell.Offset(-1, 0) + 1
            ActiveCell.Offset(0, 1).Value = cbotype.Value
            ActiveCell.Offset(0, 2).Value = txtTestCaseId
            ActiveCell.Offset(0, 3).Value = txtTestCaseName
            ActiveCell.Offset(0, 4).Value = Now
            ActiveCell.Offset(0, 5).Value = Environ("username")

Set WB = Workbooks.Add
        WB.SaveAs "C:\Users\" & Environ("USERNAME") _
        & "\Desktop\VBA Programming\Test Case\" & txtTestCaseId & "-" & txtTestCaseName

range("B2").Value = "Test Case Id"
range("B2").Font.Size = 12
range("B2").Font.Bold = True

range("C2").Value = txtTestCaseId
range("C2").Font.Size = 12
range("C2").Font.Bold = True

range("B3").Value = "Test Case Name"
range("B3").Font.Size = 12
range("B3").Font.Bold = True

range("C3").Value = txtTestCaseName
range("C3").Font.Size = 12
range("C3").Font.Bold = True

range("B4").Value = "MNS Engineer Client Version"
range("B4").Font.Size = 12

range("B5").Value = "Server IP"
range("B5").Font.Size = 10

range("C5").Value = txtServerIp
range("C5").Font.Size = 10

range("D5").Value = "Instance"
range("D5").Font.Size = 10

range("E5").Value = txtInstance
range("E5").Font.Size = 10

range("B6").Value = "ME Issue Number"
range("B6").Font.Size = 10

range("C6").Value = txtMeIssueNumber
range("C6").Font.Size = 10
range("C6").Font.Bold = True

range("B7").Value = "SR Number"
range("B7").Font.Size = 10

range("C7").Value = txtSrNumber
range("C7").Font.Size = 10
range("C7").Font.Bold = True

range("B8").Value = "Country/LBU"
range("B8").Font.Size = 10

range("C8").Value = txtLbu
range("C8").Font.Size = 10
range("C8").Font.Bold = True

range("B9").Value = "Login Credentials"
range("B9").Font.Size = 10

range("B10").Value = "Project Name"
range("B10").Font.Size = 10
range("B10").Font.Bold = True

range("B11").Value = "SL No"
range("B11").Font.Size = 10
range("B11").Font.Bold = True

range("C11").Value = "Test Case Description"
range("C11").Font.Size = 10
range("C11").Font.Bold = True

range("D11").Value = "Expected Result"
range("D11").Font.Size = 10
range("D11").Font.Bold = True

range("E11").Value = "Result"
range("E11").Font.Size = 10
range("E11").Font.Bold = True

range("F11").Value = "Comments"
range("F11").Font.Size = 10
range("F11").Font.Bold = True
Columns.AutoFit
range("B2:F20").Borders.LineStyle = xlContinuous
WB.Save
End Sub



