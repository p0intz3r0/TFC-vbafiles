Sub rct_layout(valClient, valCollab, valDateF, valDelaiR, valDatePoste, valFonction, valLibelle, valSba, valIsForfait, valSommeHT, valSommeTTC, txfac, valLibelle2, newnumber, valisavoir)
Dim gen As Worksheet
Set gen = ActiveWorkbook.Sheets("GENERATEUR RCT")
Dim clsClient As clientClass
Set clsClient = New clientClass
clsClient.clt_nom = valClient
clsClient.clt_id = get_client_number(clsClient.clt_nom)
clsClient.get_client_info clsClient
gen.Range("J13").Value = valDateF
gen.Range("J14").Value = newnumber
gen.Range("J15").Value = clsClient.clt_secteur
gen.Range("J16").Value = clsClient.clt_id
gen.Range("I19").Value = clsClient.clt_nom
gen.Range("I20").Value = clsClient.adr1
gen.Range("i21").Value = clsClient.adr2
gen.Range("I22").Value = clsClient.adr3
gen.Range("I23").Value = clsClient.adr4
gen.Range("I24").Value = clsClient.adr5
gen.Range("i25").Value = clsClient.clt_tva
gen.Range("a33").Value = valLibelle
gen.Range("a37").Value = valDatePoste
gen.Range("C33").Value = valFonction
gen.Range("C37").Value = valCollab
If valIsForfait = True Then
gen.Range("f33").Value = vbNullString
gen.Range("g33").Value = vbNullString
gen.Range("h33").Value = valSommeHT
gen.Range("j33").Value = valSommeHT
gen.Range("i33").Value = "="
Else
gen.Range("f33").Value = txfac * 0.01
gen.Range("g33").Value = "x"
gen.Range("h33").Value = valSba
gen.Range("j33").Value = valSommeHT
gen.Range("i33").Value = "="
End If
gen.Range("J39").Value = valSommeHT
If valSommeHT <> valSommeTTC Then
gen.Range("H41").Value = "20% TVA"
gen.Range("J41").Value = valSommeTTC - valSommeHT
Else
gen.Range("H41").Value = vbNullString
gen.Range("J41").Value = vbNullString
End If
gen.Range("j43").Value = valSommeTTC
gen.Range("C47").Value = valDelaiR
gen.Range("h47").Value = valDelaiR + valDateF
gen.Range("J50").Value = newnumber

If valisavoir = True Then
gen.Range("A13").Value = "AVOIR"
Else
gen.Range("A13").Value = "FACTURE"
End If
   If clsClient.clt_factor > 0 Then
        gen.Range("a53").Value = Sheets("BDD VBA").Range("K1")
    Else
        gen.Range("a53").Value = Sheets("BDD VBA").Range("a1")
    End If
    Sheets("GENERATEUR RCT").Activate
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        factpath & newnumber & ".pdf"
        MsgBox "Facture enregistrée en PDF et dans la BDD"
        End With
End Sub
