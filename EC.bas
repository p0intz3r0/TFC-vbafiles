Option Explicit
Sub EC()
    Dim wsEC As Worksheet
    Dim dDate As Date
    Dim sJournal As String
    Dim strLibelle As String
    Dim dMontantHT As Double
    Dim dMontantTTC As Double
    Dim dDateEch As Date
    Dim iNum As String
    Dim strJournal As String
    Dim strFactType As String
    Dim iLastRow As Integer
    Dim strClient As String
    Dim start As Long
    Dim access_rs As DAO.Recordset
    Dim access_db As DAO.Database
    Set access_db = DAO.DBEngine.OpenDatabase(db_path)
    start = 0
    Set wsEC = ActiveWorkbook.Sheets("EC")
    iLastRow = wsEC.Cells(Rows.Count, "B").End(xlUp).Row
    Dim sqlCMD As String
    sqlCMD = "SELECT * FROM [FACT] WHERE NUMFACTURE>" & start & ";"
    iLastRow = wsEC.Cells(Rows.Count, "B").End(xlUp).Row
    Set access_rs = access_db.OpenRecordset(sqlCMD)
    If Not (access_rs.EOF And access_rs.BOF) Then
        access_rs.MoveFirst
        Do
            iLastRow = iLastRow + 1
            Debug.Print strClient
          '  If access_rs.Fields(1) = "FRCT" Then
            If Left(access_rs.Fields(1), 1) = "F" Then
            strLibelle = access_rs.Fields(0)
            dDate = access_rs.Fields(4)
            '  dDate = Format(dDate, "dd/mm/yy")
            dMontantHT = Abs(access_rs.Fields(9))
            dMontantTTC = Abs(access_rs.Fields(10))
            If IsNull(access_rs.Fields(14)) Then
            dDateEch = dDate + get_client_delai_from_num(access_rs.Fields(3))
            Else
            dDateEch = dDate + access_rs.Fields(14)
            End If
                strClient = access_rs.Fields(3)
               ' strClient = Replace(strClient, " ", vbNullString)  'DEBUG'
              '  strClient = Replace(strClient, Chr(45), "") 'STRING OPERATIONS FOR STR_CLIENT '
              '  strClient = Replace(strClient, Chr(39), "") 'WE ARE NOW USING THE CLIENT NUM '
              '  strClient = UCase("C" + Left(strClient, 11))
                wsEC.Cells(iLastRow, 1) = strClient
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 5) = dMontantTTC
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 70660400
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 6) = dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 44571200
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 6) = dMontantTTC - dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
            ElseIf access_rs.Fields(1) = "A" Then
                strLibelle = access_rs.Fields(0)
            dDate = access_rs.Fields(4)
            '  dDate = Format(dDate, "dd/mm/yy")
            dMontantHT = Abs(access_rs.Fields(9))
            dMontantTTC = Abs(access_rs.Fields(10))
            If IsNull(access_rs.Fields(14)) Then
            dDateEch = dDate + get_client_delai_from_num(access_rs.Fields(3))
            Else
            dDateEch = dDate + access_rs.Fields(14)
            End If
             ''    strClient = get_client_name2(access_rs.Fields(3))
             '   strClient = Replace(strClient, " ", vbNullString)
              '  strClient = Replace(strClient, Chr(45), "")
             '   strClient = Replace(strClient, Chr(39), "")
             '   strClient = UCase("C" + Left(strClient, 11))
                wsEC.Cells(iLastRow, 1) = strClient
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 6) = dMontantTTC
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 70660400
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 5) = dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 44571200
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 5) = dMontantTTC - dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
            Else
            End If
            access_rs.MoveNext
        Loop Until access_rs.EOF
    End If
End Sub
Sub EC_rct()
    Dim wsEC As Worksheet
    Dim dDate As Date
    Dim sJournal As String
    Dim strLibelle As String
    Dim dMontantHT As Double
    Dim dMontantTTC As Double
    Dim dDateEch As Date
    Dim iNum As String
    Dim wsDB As Worksheet
    Dim rStrt As Range
    Dim strJournal As String
    Dim x As Double
    Dim y As Double
    Dim strFactType As String
    Dim iLastRow As Integer
    Dim strClient As String
    Dim i As Range
    Dim access_rs As DAO.Recordset
    Dim access_db As DAO.Database
    Set access_db = DAO.DBEngine.OpenDatabase(db_path)
    Dim start As Long
    Dim sqlCMD As String
    sqlCMD = "SELECT * FROM [FACTRCT] WHERE RCT_ID >" & start & ";"
    Set wsEC = ActiveWorkbook.Sheets("EC")
    iLastRow = wsEC.Cells(Rows.Count, "B").End(xlUp).Row
    Set access_rs = access_db.OpenRecordset(sqlCMD)
    If Not (access_rs.EOF And access_rs.BOF) Then
        access_rs.MoveFirst
        Do
            strClient = access_rs.Fields(3)
            strLibelle = access_rs.Fields(0)
            dDate = access_rs.Fields(4)
            '  dDate = Format(dDate, "dd/mm/yy")
            dMontantHT = Abs(access_rs.Fields(11))
            dMontantTTC = Abs(access_rs.Fields(12))
            dDateEch = dDate + access_rs.Fields(10)
           ' strClient = Replace(strClient, " ", vbNullString)
           ' strClient = Replace(strClient, Chr(45), "")
           ' strClient = Replace(strClient, Chr(39), "")
          '  strClient = UCase("C" + Left(strClient, 11))
           ' Debug.Print strClient
            If Left(access_rs.Fields(1), 1) = "F" Then
                wsEC.Cells(iLastRow, 1) = strClient
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 5) = dMontantTTC
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 70660400
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 6) = dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 44571200
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 6) = dMontantTTC - dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
            ElseIf Left(access_rs.Fields(1), 1) = "A" Then
                wsEC.Cells(iLastRow, 1) = strClient
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 6) = dMontantTTC
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 70660400
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 5) = dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
                iLastRow = iLastRow + 1
                wsEC.Cells(iLastRow, 1) = 44571200
                wsEC.Cells(iLastRow, 2) = dDate
                wsEC.Cells(iLastRow, 3) = "VE"
                wsEC.Cells(iLastRow, 4) = strLibelle
                wsEC.Cells(iLastRow, 5) = dMontantTTC - dMontantHT
                wsEC.Cells(iLastRow, 7) = dDateEch
                wsEC.Cells(iLastRow, 8) = iNum
            Else
            
            End If
            access_rs.MoveNext
        Loop Until access_rs.EOF
    End If
End Sub


