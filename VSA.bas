Option Explicit
Sub exp_vsa()
    Application.Calculation = xlCalculationManual
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Dim vsa_WS As Worksheet
    Dim sqlcmd_STR As String
    Set access_db = DAO.OpenDatabase(db_path)
    Dim i As Long
    Dim lastrow As Double
    sqlcmd_STR = "SELECT * FROM [FACT] WHERE TYPE = 'F' OR TYPE = 'A'"
    Set access_rs = access_db.OpenRecordset(sqlcmd_STR)
    Set vsa_WS = ActiveWorkbook.Sheets("VSA")
    access_rs.MoveFirst
    i = 2
    If Not (access_rs.BOF And access_rs.EOF) Then
        Do
            vsa_WS.Cells(i, 1) = "COOPTALIS"
            vsa_WS.Cells(i, 3) = access_rs.Fields(0)
            vsa_WS.Cells(i, 2) = access_rs.Fields(1)
            vsa_WS.Cells(i, 14) = access_rs.Fields(2)
            vsa_WS.Cells(i, 4) = access_rs.Fields(3)
            vsa_WS.Cells(i, 6) = access_rs.Fields(4)
            vsa_WS.Cells(i, 8) = access_rs.Fields(9)
            vsa_WS.Cells(i, 18) = access_rs.Fields(7)
            vsa_WS.Cells(i, 19) = access_rs.Fields(8)
            vsa_WS.Cells(i, 17) = access_rs.Fields(6)
            If access_rs.Fields(1) = "A" And Not IsNull(access_rs.Fields(13)) Then
                If Len(onlyDigits(access_rs.Fields(13))) > 6 Then
                    Dim temp_Srch_Index As Byte
                    temp_Srch_Index = InStr(access_rs.Fields(13), "F")
                    vsa_WS.Cells(i, 5) = Mid(access_rs.Fields(13), temp_Srch_Index + 1, 5)
                Else
            
                    vsa_WS.Cells(i, 5) = onlyDigits(access_rs.Fields(13))
                End If
            End If
            vsa_WS.Cells(i, 9) = Abs(Round(1 - (access_rs.Fields(10) / access_rs.Fields(9)), 1))
            vsa_WS.Cells(i, 10) = "EUR"
            vsa_WS.Cells(i, 12) = 1
            i = i + 1
            access_rs.MoveNext
        Loop Until access_rs.EOF
            Application.Calculation = xlCalculationAutomatic
    End If
End Sub


Function onlyDigits(s As String) As String
    Dim retval As String    
    Dim i As Integer                    
    retval = vbNullString                                 
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next       
    onlyDigits = retval
End Function


