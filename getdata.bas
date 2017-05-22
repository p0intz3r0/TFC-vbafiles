Public Function miseenpageIT(tjm, client, datefact, joursfactu, collabname, Libelle)
    Dim isfactor As Byte, facturnumber As Double, typefact As String, ato As String, totalht As Single, totalttc As Single, delairglt As Integer, gen As Worksheet
    ato = "ATO 09/16"
    Sheets("GENERATEUR ATO").Activate
    ActiveSheet.Range("I19") = client
    ActiveSheet.Range("J14") = num
    ActiveSheet.Range("J13") = datefact
    ActiveSheet.Range("D34") = collabname
    ActiveSheet.Range("A34") = ato
    ActiveSheet.Range("F34") = joursfactu
    ActiveSheet.Range("C48") = collabname
    ActiveSheet.Range("H34") = tjm
    ActiveSheet.Range("J15") = "IT"
    ActiveSheet.Range("A36") = Libelle
    facturnumber = get_last_invoice_num()
    Sheets("BDD VBA").Range("K5") = facturnumber
    ActiveSheet.Range("J14") = facturnumber
'-- RechercheV pour coordonnées client --'
    ActiveSheet.Range("I20") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 2, 0)
    ActiveSheet.Range("I21") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 3, 0)
    ActiveSheet.Range("I22") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 4, 0)
    ActiveSheet.Range("I23") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 5, 0)
    ActiveSheet.Range("I24") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 6, 0)
    ActiveSheet.Range("I25") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 9, 0)
    ActiveSheet.Range("C48") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 10, 0)
    ActiveSheet.Range("J16") = Application.VLookup(client, Sheets("BDD Clients").Range("B5:N100"), 8, 0)
    sngclient = Application.VLookup(client, Sheets("BDD Clients").Range("B5:N100"), 8, 0)
    isfactor = Application.VLookup(client, Sheets("BDD Clients").Range("B5:M100"), 11, 0)
    totalht = ActiveSheet.Range("J40").Value
    totalttc = ActiveSheet.Range("J44").Value
    delairglt = ActiveSheet.Range("C48").Value
   If isfactor = 1 Then
   ' typefact = "Facture Factor CIC"
    ElseIf isfactor = 2 Then
    ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("K1")
     typefact = "Facture Factor NATIXIS"
    Call fonctionsIT.csvnatixis(totalht, totalttc, datefact, facturnumber, delairglt, client, typefact)
       Call fonctionsIT.reportsurDBIT(tjm, sngclient, datefact, joursfactu, collabname, facturnumber, typefact, totalht, totalttc, delairglt, ato, isfactor, Libelle)
       Call fonctionsIT.exportEnPdfIT(facturnumber)
    Else
        ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("A1")
        typefact = "Facture Directe"
         ActiveSheet.Range("A53") = ""
            Call fonctionsIT.reportsurDBIT(tjm, client, datefact, joursfactu, collabname, facturnumber, typefact, totalht, totalttc, delairglt, ato, isfactor, Libelle)
       Call fonctionsIT.exportEnPdfIT(facturnumber)
End If
End Function
 Public Function reportsurDBIT(tjm, sngclient, datefact, joursfactu, collabname, facturnumber, typefact, totalht, totalttc, delairglt, ato, isfactor, Libelle)
 Dim access_db As Object
 Set access_db = DAO.DBEngine.OpenDatabase(db_path)
 Dim sql_insert As String
 Dim sql_log As String
 Dim us3r As String
 Dim strIsfactor As String
 Dim sglClient As Single
 Dim sglCollab As Single
 Dim tpCollabName As String
 tpCollabName = Split(collab, " ")(0)
 sglCollab = get_collab_number(tpCollabName)
 If isfactor = 1 Or 2 Then
 strIsfactor = "F"
 Else
 strIsfactor = "N"
 End If
 typefact = Left(typefact, 1)
 'User = nom d'utilisateur de session Windows '
 'MAJ 01-08-16 : Au lieu de reporter les informations sur une feuille, elles sont envoyées dans la BDD '
 'MAJ 31-08-16 : TODO : Remplacer client et collab par leurs ID respectifs '
    sql_insert = "INSERT INTO [FACT] (NUMFACTURE, TYPE, COLLAB,CLIENT, DATEFAC, PERIODE, TJM,  LIBELLE, NBJOURS, MONTANTHT, MONTANTTTC, REGLEMENT, LIBELLE2, COLLAB_ID) " & _
    "VALUES (" & facturnumber & "," & Chr(39) & typefact & Chr(39) & "," & Chr(39) & collabname & Chr(39) & ", " & Chr(39) & sngclient & Chr(39) & "," & Chr(39) & datefact & Chr(39) & "," & _
     Chr(39) & Month(datefact) & Chr(39) & "," & Chr(39) & tjm & Chr(39) & ", " & Chr(39) & ato & Chr(39) & "," & Chr(39) & joursfactu & Chr(39) & "," & Chr(39) & totalht & _
     Chr(39) & "," & Chr(39) & totalttc & Chr(39) & ", " & Chr(39) & strIsfactor & Chr(39) & "," & Chr(39) & Libelle & Chr(39) & sglCollab & Chr(39) & ");"
    Debug.Print sql_insert
access_db.Execute (sql_insert)
'La deuxième commande sert à logger les actions dans une table separée, pour plus de tracabilité '
    sql_log = "INSERT INTO [LOG] (username, timest, command, num) VALUES" & _
    "(" & Chr(39) & us3r & Chr(39) & "," & Chr(39) & Now() & Chr(39) & ", " & Chr(39) & Left(sql_insert, 12) & facturnumber & Chr(39) & "," & facturnumber & ");"
    Debug.Print sql_log
access_db.Execute (sql_log)
access_db.Close
End Function
Public Function exportEnPdfIT(facturnumber)
    Sheets("GENERATEUR ATO").Activate
    facturnumber = CStr(facturnumber)
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\FACTURES 2016\" & facturnumber & ".pdf"
        ',Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False, DisplayAlerts:=True
    End With
End Function
Public Function csvnatixis(totalht, totalttc, datefact, facturnumber, delairglt, client, typefact)
Dim btypefact As String
btypefact = Left(typefact, 1)
Dim iNumeroClient As Long
iNumeroClient = client
Dim lastrow As Long
Dim WsCsv As Worksheet
Set WsCsv = Sheets("CSVNATIXIS")
WsCsv.Activate
ptrlastrow = WsCsv.Cells(Rows.Count, "B").End(xlUp).Row
ptrlastrow = ptrlastrow + 1
WsCsv.Range("A" & ptrlastrow).Value = btypefact
WsCsv.Range("B" & ptrlastrow).Value = facturnumber
WsCsv.Range("C" & ptrlastrow).Value = datefact
WsCsv.Range("D" & ptrlastrow).Value = iNumeroClient
WsCsv.Range("E" & ptrlastrow).Value = totalht
WsCsv.Range("F" & ptrlastrow).Value = totalttc
WsCsv.Range("G" & ptrlastrow) = delairglt
Dim dEcheance As Date
dEcheance = datefact + delairglt
WsCsv.Range("H" & ptrlastrow) = dEcheance
WsCsv.Range("I" & ptrlastrow) = "VIR"
End Function
