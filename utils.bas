Option Explicit
Sub show_userform1()
    Load UserForm1
    UserForm1.Show
End Sub
'---------------------------------------------------------------------------------------
' Method : test_customclass
' Author : p0intz3r0
' Date   : 27/09/2016
' Purpose: For debugging purposes only
'---------------------------------------------------------------------------------------
Sub test()
' Call utils.save_natixis
Debug.Print Application.OperatingSystem
End Sub
Sub mailShow()
    mailForm.Show
End Sub
Public Function frais_layout(fr_num As Double, fr_clt As String, fr_mtn1 As Double, fr_mtn2 As Double, fr_ht As Double, fr_ttc As Double, fr_delai As Double, fr_date As Date, fr_libelle As String)
    Dim clsClient As clientClass
    Set clsClient = New clientClass
    Dim gen As Worksheet
    Set gen = ActiveWorkbook.Sheets("GENERATEUR FRAIS")
    clsClient.clt_nom = fr_clt
    clsClient.clt_id = get_client_number(clsClient.clt_nom)
    ' clientnb = get_client_number(newinvoice.client)
    clsClient.get_client_info clsClient
    gen.Range("J13").Value = fr_date
    gen.Range("i18").Value = clsClient.clt_nom
    gen.Range("J14").Value = fr_num
    gen.Range("J15").Value = clsClient.clt_secteur
    gen.Range("J16").Value = clsClient.clt_id
    gen.Range("I19").Value = clsClient.adr1
    gen.Range("I20").Value = clsClient.adr2
    gen.Range("I21").Value = clsClient.adr3
    gen.Range("I22").Value = clsClient.adr4
    gen.Range("I23").Value = clsClient.adr5
    gen.Range("I25").Value = clsClient.clt_tva
    gen.Range("A34").Value = fr_libelle
    gen.Range("H34").Value = fr_mtn1
    gen.Range("J34").Value = fr_mtn1
    gen.Range("H36").Value = fr_mtn2
    gen.Range("J36").Value = fr_mtn2
    gen.Range("J40").Value = fr_ht
    gen.Range("J44").Value = fr_ttc
    gen.Range("J42").Value = fr_ttc - fr_ht
    If Not IsNull(fr_delai) Then
        gen.Range("C47").Value = fr_delai
    Else
        gen.Range("C47").Value = clsClient.clt_delai
    End If
    gen.Range("h47").Value = fr_date + gen.Range("C47").Value
    gen.Range("J51").Value = fr_num
    If clsClient.clt_factor > 0 Then
        gen.Range("a54").Value = Sheets("BDD VBA").Range("K1")
    Else
        gen.Range("a54").Value = Sheets("BDD VBA").Range("a1")
    End If
    Dim fr_id As Double
    fr_id = fr_num
    Call send_to_RCTFrais(fr_id, fr_clt, fr_mtn1, fr_mtn2, fr_ht, fr_ttc, fr_delai, fr_date, fr_libelle)
    Call exportEnPdf_frais(fr_id)
End Function

Public Function send_to_RCTFrais(fr_id, fr_clt, fr_mtn1, fr_mtn2, fr_ht, fr_ttc, fr_delai, fr_date, fr_libelle)
    Dim access_db As DAO.Database
    Dim newnumber As Double
    newnumber = get_last_invoice_num()
    Set access_db = DAO.OpenDatabase(db_path)
    Dim is_factor_query_result As Boolean
    Dim sql_command As String
    Dim typ As String
    sql_command = "INSERT INTO [FRAIS] (FRAIS_ID, FRAIS_CLT, FRAIS_MTN1, FRAIS_MTN2, FRAIS_THT, FRAIS_TTC, FRAIS_DELAI, FRAIS_DATE, FRAIS_LIBELLE ) VALUES ('" & fr_id & "','" & fr_clt & "', '" & fr_mtn1 & "', '" & fr_mtn2 & "', '" & fr_ht & "', '" & fr_ttc & _
        "', '" & fr_delai & "', '" & fr_date & "', '" & fr_libelle & "');"
    Debug.Print sql_command
    access_db.Execute (sql_command)
    typ = "FRCT"
    sql_command = "INSERT INTO [FACT] (NUMFACTURE, TYPE, LIBELLE, CLIENT, MONTANTTTC ) VALUES ('" & fr_id & _
        "', 'FFRAIS','" & fr_libelle & "','" & fr_clt & "','" & fr_ttc & "');"
    access_db.Execute (sql_command)
    access_db.Close
End Function
Public Function exportEnPdf_frais(facturnumber)
    Sheets("GENERATEUR FRAIS").Activate
    facturnumber = CStr(facturnumber)
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
            "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2017\FACTURES 2017\" & facturnumber & ".pdf"
        ',Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
         :=False, OpenAfterPublish:=False, DisplayAlerts:=True
    End With
End Function

Public Function sales_estimates(ByVal ws As Worksheet)
    Dim strt_rng As Range
    Set strt_rng = ws.Range("D2:D1500")
    Dim cll As Range
    Dim tp_sum As Double
    For Each cll In strt_rng
        If Not IsEmpty(cll) Then
            tp_sum = tp_sum + ws.Cells(cll.Row, cll.Column + 13).Value * ws.Cells(cll.Row, cll.Column + 7).Value
            Debug.Print tp_sum
        End If
    Next cll
    sales_estimates = tp_sum
End Function

Sub show_batch()
BatchForm.Show
End Sub
Public Function save_natixis()
Dim svWorksheet As Worksheet
Set svWorksheet = ActiveWorkbook.Sheets("CSVNATIXIS")
'fileformat:=51 : fichier .xls '
'svWorksheet.SaveAs(factpath & "REMFACTO", xlWorkbookNormal)
End Function
Public Function clear_natixis()
ActiveWorkbook.Sheets("CSVNATIXIS").Range("A2:ZZ1000").Clear
End Function
