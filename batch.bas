Option Explicit
Sub main(ByVal datefact As Date)
    Application.ScreenUpdating = False
    Dim isfac As Range
    Dim x, y, n As Integer
    Dim tjm, numfa As Double
    Dim collabname, client As String
    Dim joursfactu As Double
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim i As Variant
    Dim firstnumfa As Double
    Dim cell As Range
    Dim Libelle As String
    Dim tclient As String
    Dim tmpadrr As String
    Sheets("BDD Collabs").Activate
    Call utils.clear_natixis
    Select Case MsgBox("Etes vous sur de vouloir éditer toutes les factures IT ? ", vbOKCancel Or vbExclamation, Application.Name)
        Case vbCancel
            Exit Sub
        Case vbOK
            StartTime = Timer
            Set isfac = Sheets("BDD Collabs").Range("S2:S1500")
            For Each i In isfac
                If i = 1 Then
                    n = n + 1
                    i.Select
                    x = i.Row
                    y = i.Column
                    numfa = get_last_invoice_num()
                    Cells(x, y + 2).Value = numfa
                    joursfactu = Cells(x, y - 2).Value
                    Dim newinvoice As Variant
                    Set newinvoice = New InvoiceClass
                    If joursfactu < 0 Then
                        newinvoice.isavoir = True
                    Else
                        newinvoice.isavoir = False
                    End If
                    Libelle = Cells(x, y - 1).Value
                    tjm = Cells(x, y - 8).Value
                    tclient = Cells(x, y - 13).Value
                    ' newinvoice.Libelle2 = Cells(x, y - 5).Value
                    If Left(tclient, 5) = "OPEN " Then
                        client = Left(tclient, 5)
                        newinvoice.Libelle2 = Mid(Libelle, 10) & " Centre de " & Mid(tclient, 5)
                        newinvoice.Libelle = Left(Libelle, 10)
                    ElseIf Left(tclient, 5) = "ATOS " Or Left(tclient, 5) = "BULL " Then
                        client = tclient
                        newinvoice.Libelle2 = Cells(x, y - 6).Value
                        newinvoice.Libelle = Left(Libelle, 10)
                        tmpadrr = Cells(x, y - 5).Value
                        newinvoice.adresselivr = tmpadrr
                    ElseIf Left(tclient, 5) = "MODIS" Then
                        client = tclient
                        newinvoice.Libelle2 = "Centre :" & Cells(x, y - 6).Value & " "
                        newinvoice.Libelle = Left(Libelle, 10)
                    Else
                        client = tclient
                        newinvoice.Libelle = Left(Libelle, 10)
                        newinvoice.Libelle2 = Mid(Libelle, 10)
                    End If
                    collabname = Cells(x, y - 15).Value
                    newinvoice.tjm = tjm
                    newinvoice.client = client
                    newinvoice.joursfact = joursfactu
                    newinvoice.delairglt = get_client_delai(client)
                    newinvoice.collab = collabname
                    newinvoice.invoicedate = datefact
                    newinvoice.send_to_db newinvoice
                    newinvoice.new_invoice_layout newinvoice
                    newinvoice.new_invoice_pdf_save newinvoice
                    Sheets("BDD Collabs").Activate
                End If
            Next i
            Application.ScreenUpdating = True
    End Select
    Call utils.save_natixis
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "Edité " & n & " factures en " & SecondsElapsed & " secondes"
End Sub

Sub batch_test(ByVal datefact As Date)

   On Error GoTo batch_test_Error

    Application.ScreenUpdating = False
    Dim isfac As Range
    Dim x, y, n As Integer
    Dim tjm, numfa As Double
    Dim collabname, client As String
    Dim joursfactu As Double
    Dim i As Variant
    Dim firstnumfa As Double
    Dim cell As Range
    Dim Libelle As String
    Dim tclient As String
    Sheets("BDD Collabs").Activate
    'Call utils.clear_natixis
    '  Select Case MsgBox("Etes vous sur de vouloir éditer toutes les factures IT ? ", vbOKCancel Or vbExclamation, Application.Name)
    ' Case vbCancel
    '    Exit Sub
    ' Case vbOK
    Set isfac = Sheets("BDD Collabs").Range("S2:S1500")
    For Each i In isfac
        If i = 1 Then
            n = n + 1
            i.Select
            x = i.Row
            y = i.Column
            numfa = get_last_invoice_num()
            Cells(x, y + 2).Value = numfa
            joursfactu = Cells(x, y - 2).Value
            Dim newinvoice As Variant
            Set newinvoice = New InvoiceClass
            If joursfactu < 0 Then
                newinvoice.isavoir = True
            Else
                newinvoice.isavoir = False
            End If
            Libelle = Cells(x, y - 1).Value
            tjm = Cells(x, y - 8).Value
            tclient = Cells(x, y - 13).Value
            ' newinvoice.Libelle2 = Cells(x, y - 5).Value
            If Left(tclient, 5) = "OPEN " Then
                client = Left(tclient, 5)
                newinvoice.Libelle2 = Mid(Libelle, 10) & " Centre de " & Mid(tclient, 5)
                newinvoice.Libelle = Left(Libelle, 10)
            ElseIf Left(tclient, 5) = "ATOS " Or Left(tclient, 5) = "BULL " Then
                client = tclient
                newinvoice.Libelle2 = Cells(x, y - 6).Value
                newinvoice.Libelle = Left(Libelle, 10)
            ElseIf Left(tclient, 5) = "MODIS" Then
                client = tclient
                newinvoice.Libelle2 = "Centre :" & Cells(x, y - 6).Value & " "
                newinvoice.Libelle = Left(Libelle, 10)
            Else
                client = tclient
                newinvoice.Libelle = Left(Libelle, 10)
                newinvoice.Libelle2 = Mid(Libelle, 10)
            End If
            collabname = Cells(x, y - 15).Value
            newinvoice.tjm = tjm
            newinvoice.client = client
            newinvoice.joursfact = joursfactu
            newinvoice.delairglt = get_client_delai(client)
            newinvoice.collab = collabname
            newinvoice.invoicedate = datefact
            'newinvoice.send_to_db newinvoice
            ' newinvoice.new_invoice_layout newinvoice
            '  newinvoice.new_invoice_pdf_save newinvoice
            Sheets("BDD Collabs").Activate
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "Tests OK "
    '  End Select
    ' Call utils.save_natixis

   On Error GoTo 0
   Exit Sub

batch_test_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure batch_test of Sub batch"

End Sub
