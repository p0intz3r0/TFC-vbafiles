Option Explicit
Public Function avoirtotal(orclient, orjourspr, orTJM, ordate, ordelai, orcollab, ornum, compl)
  '  ato = "ATO 01/16"
  Dim isfactor As Byte
  Dim totalht, totalttc, delairglt, facturnumber As Double
  Dim typeavoir As String
    Sheets("GENERATEUR AVOIR").Activate
    ActiveSheet.Range("I19") = orclient
    ActiveSheet.Range("J13") = ordate
    ActiveSheet.Range("D34") = orcollab
    ActiveSheet.Range("F34") = -orjourspr
    ActiveSheet.Range("C48") = orcollab
     ActiveSheet.Range("H34") = orTJM
     ActiveSheet.Range("J15") = "IT"
     ActiveSheet.Range("A34") = "Avoir sur facture n°" & ornum
         ActiveSheet.Range("A35") = compl
    facturnumber = Sheets("Facturation ATO 2016").Range("B2")
    facturnumber = facturnumber + 1
    Sheets("Facturation ATO 2016").Range("B2") = facturnumber
        ActiveSheet.Range("J14") = facturnumber
'-- RechercheV pour coordonnées orclient --'
    ActiveSheet.Range("I20") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 2, 0)
    ActiveSheet.Range("I21") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 3, 0)
    ActiveSheet.Range("I22") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 4, 0)
    ActiveSheet.Range("I23") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 5, 0)
    ActiveSheet.Range("I24") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 6, 0)
    ActiveSheet.Range("I25") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 9, 0)
    ActiveSheet.Range("C48") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 10, 0)
    ActiveSheet.Range("J16") = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:N100"), 8, 0)
    isfactor = Application.VLookup(orclient, Sheets("BDD Clients").Range("B5:M100"), 11, 0)
    totalht = ActiveSheet.Range("J40").Value
    totalttc = ActiveSheet.Range("J44").Value
    delairglt = ActiveSheet.Range("C48").Value
   If isfactor = 1 Then
    ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("A5")
    typeavoir = "Avoir Factor"
    ElseIf isfactor = 2 Then
    
    ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("K1")
    typeavoir = "Avoir Factor"
    Else
        ActiveSheet.Range("A54") = Sheets("BDD VBA").Range("A1")
        typeavoir = "Avoir Direct"
End If
    Call reportsurDBavoir(orTJM, orclient, facturnumber, orcollab, totalht, totalttc, delairglt, ordate, ornum, orjourspr, typeavoir)
    Call exportEnPdfavoir(facturnumber)
    MsgBox " Enregistrement de l'avoir " & facturnumber & "réussi"
End Function

Public Function reportsurDBavoir(orTJM, orclient, facturnumber, orcollab, totalht, totalttc, delairglt, ordate, ornum, orjourspr, typeavoir)
    Dim lastrow As Integer
    lastrow = Sheets("Facturation ATO 2016").Cells(Rows.Count, "B").End(xlUp).Row
    lastrow = lastrow + 1
    Sheets("Facturation ATO 2016").Range("B" & lastrow) = facturnumber
    Sheets("Facturation ATO 2016").Range("C" & lastrow) = typeavoir
    Sheets("Facturation ATO 2016").Range("D" & lastrow) = orclient
    Sheets("Facturation ATO 2016").Range("E" & lastrow) = "Avoir sur facture n°" & ornum
    Sheets("Facturation ATO 2016").Range("F" & lastrow) = orcollab
    Sheets("Facturation ATO 2016").Range("G" & lastrow) = -orjourspr
    Sheets("Facturation ATO 2016").Range("H" & lastrow) = orTJM
    Sheets("Facturation ATO 2016").Range("I" & lastrow) = ordate
    Sheets("Facturation ATO 2016").Range("J" & lastrow) = totalht
    Sheets("Facturation ATO 2016").Range("K" & lastrow) = totalttc
    Sheets("Facturation ATO 2016").Range("L" & lastrow) = delairglt
End Function
Public Function exportEnPdfavoir(facturnumber)
    Sheets("GENERATEUR AVOIR").Activate
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
        "J:\1 - Contrôle de Gestion\2 - Facturation Client\Facturation 2016\temp" & facturnumber & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
    End With
End Function

