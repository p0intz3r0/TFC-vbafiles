'---------------------------------------------------------------------------------------
' File   : data_ops
' Author : p0intz3r0
' Date   : 26/09/2016
' Purpose: Multiple data handling functions
'---------------------------------------------------------------------------------------
Option Explicit
'+--------------------------------------------------------------------+
'|  Diagramme algorithmique fonction get_all_info                     |
'|  @params :oldinvoice (object, type = newinvoice, val=num)          |
'|  @returns : oldinVoice[tjm,client,collab,joursfact]                |
'+--------------------------------------------------------------------+
'|  If typefact =    F    |                 FRCT            |   A     |
'+--------------------------------------------------------------------+
'|          |             |FORFAIT                          |    |    |
'|          |             |   O                     N       |    |    |
'|          |             +---------------------------------+    |    |
'|          v             |   v           |         v       |    |    |
'+----------------------------------------------------------+    |    |
'|                        |               |                 |    |    |
'|With oldinvoice:        |               |                 |    |    |
'|    .collab  = collab   |.collab        |.collab          |    |    |
'|    .tjm = TJM          |=Candidat      |=Candidat        |    |    |
'|    .joursfact = jours  |.TJM           |.TJM = SBA       |    |    |
'|    .client = client    |= Forfait      |.jours =         |    |    |
'|                        |.jours = 1     |taux (%)         |    |    |
'|                        |               |                 |    |    |
'|Returns oldinvoice(all) |               |                 |    |    |
'|                        |               |                 |    v    |
'|                        |               |                 |Msgbox   |
'|                        |               |                 |return   |
'|                        |               |                 |error    |
'+--------------------------------------------------------------------+

Public Function get_all_info(oldinvoice)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Dim sql_command As String
    Dim typeFacture As String
    Set access_db = DAO.OpenDatabase(db_path)
    sql_command = "SELECT TYPE FROM [FACT] WHERE NUMFACTURE = " & oldinvoice.number '---------------'
    Set access_rs = access_db.OpenRecordset(sql_command)
    If Not (access_rs.BOF And access_rs.EOF) Or access_rs.Fields(0) = Null Then
        typeFacture = access_rs.Fields(0)
    Else
        MsgBox "Facture n° " & oldinvoice.number & " non trouvée dans la BDD. Contactez votre administrateur "
    End If
        sql_command = "SELECT COLLAB, CLIENT, DATEFAC, TJM, NBJOURS FROM [FACT] WHERE NUMFACTURE = " & oldinvoice.number & ";"
        Debug.Print sql_command
        Set access_rs = access_db.OpenRecordset(sql_command)
        If Not (access_rs.BOF And access_rs.EOF) Or access_rs.Fields(0) = Null Then
            oldinvoice.collab = access_rs.Fields(0)
            oldinvoice.client = access_rs.Fields(1)
            oldinvoice.invoicedate = access_rs.Fields(2)
            oldinvoice.tjm = access_rs.Fields(3)
            oldinvoice.joursfact = access_rs.Fields(4)
        Else
            MsgBox "Erreur. Impossible de retrouver les informations de la facture n°" & oldinvoice.number & vbNewLine & _
                "Contactez votre administrateur"
        End If
        oldinvoice.client = get_client_name(oldinvoice)
        Debug.Print sql_command
End Function

Public Function get_collab_surname(collabTP)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Dim sql_command As String
    Set access_db = DAO.OpenDatabase(db_path)
    sql_command = "SELECT PRENOM FROM [COLLAB] WHERE NOM = '" & collabTP & "';"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    If Not (access_rs.BOF And access_rs.EOF) Or access_rs.Fields(0) = Null Then
        get_collab_surname = access_rs.Fields(0)
    Else
        MsgBox "Impossible de retrouver le prénom du collaborateur"
        Exit Function
    End If
End Function
Public Function get_email_adress(nbclient)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Dim sql_command As String
    Set access_db = DAO.OpenDatabase(db_path)
    sql_command = "SELECT MAIL FROM [CLT] WHERE REFCLIENT = " & nbclient & ";"
    Set access_rs = access_db.OpenRecordset(sql_command)
    If Not (access_rs.BOF And access_rs.EOF) Or access_rs.Fields(0) = Null Then
        get_email_adress = access_rs.Fields(0)
    Else
        MsgBox "Adresse e-mail vide"
        Exit Function
    End If
End Function
Public Function create_new_client(newName, newAdr1, newAdr2, newAdr3, newAdr4, newMail, newSect, newTVA, newDelai, newNum, newfactor)
    Dim access_db As DAO.Database
    Set access_db = DAO.OpenDatabase(db_path)
    Dim is_factor_query_result As Boolean
    Dim sql_command As String
    Dim numfactor As Byte
    If newfactor = True Then
        numfactor = 2
    Else
        numfactor = 0
    End If
    sql_command = "INSERT INTO [CLT] (REFCLIENT, CLTNOM, ADRESSE1, ADRESSE2, ADRESSE3, ADRESSE4,SECTEUR, TVA, DELAI, FACTORDIRECT, MAIL) VALUES (" & _
        newNum & ",'" & newName & "','" & newAdr1 & "','" & newAdr2 & "', '" & newAdr3 & "', '" & newAdr4 & "', '" & newSect & "', '" & newTVA & "', " & newDelai & "," & numfactor & ",'" & newMail & "');"
    Debug.Print sql_command
    access_db.Execute (sql_command)
    access_db.Close
    MsgBox "Client enregistré dans la BDD"
End Function
Public Function invoices_client_from_nb(varClient, varStart)
    Dim sql_command As String
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Dim nbclient As Double
    nbclient = get_client_number(varClient)
    Set access_db = DAO.OpenDatabase(db_path)
    sql_command = "SELECT NUMFACTURE, MONTANTHT FROM [FACT] WHERE CLIENT = " & nbclient & " AND NUMFACTURE > " & varStart & ";"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    If Not (access_rs.BOF And access_rs.EOF) Then
        Dim wsnew As Worksheet
        Set wsnew = Sheets.Add
       ' wsnew.Name = "ATOS" 'Uncomment this for atos_mail() #dirty'
        wsnew.Range("A1").CopyFromRecordset access_rs
        Call mail.mail2(wsnew, nbclient) 'Comment this for atos_mail() '
    Else
        MsgBox "Erreur. Aucune correspondance trouvée"
    End If

End Function

Public Function send_to_RCTDB(valClient, valCollab, valDateF, valDelaiR, valDatePoste, valFonction, valLibelle, valSba, valIsForfait, valSommeHT, valSommeTTC, vartxfac, valLibelle2, valisavoir)
    Dim access_db As DAO.Database
    Dim newnumber As Double
    newnumber = get_last_invoice_num()
    Set access_db = DAO.OpenDatabase(db_path)
    Dim is_factor_query_result As Boolean
    Dim sql_command As String
    Dim typ As String
    If valIsForfait = True Then
        typ = "Facture Forfait"
    Else
        typ = "Facture sur SBA"
    End If
    If valisavoir = True Then
    typ = "Avoir"
    End If
    
    sql_command = "INSERT INTO [FACTRCT] (RCT_ID, RCT_TYPE, RCT_COLLAB, RCT_CLIENT, RCT_DATEFAC, RCT_LIBELLE, RCT_POSTE, RCT_ISFORFAIT, RCT_PERCENT, RCT_SALAIRE, RCT_DELAIRGLT, " & _
        "RCT_MONTANTHT, RCT_MONTANTTTC, RCT_LIBELLE2, RCT_PRISEPOSTE ) VALUES ('" & newnumber & "','" & typ & "', '" & valCollab & "', '" & valClient & "', '" & valDateF & "', '" & valLibelle & _
        "', '" & valFonction & "', '" & valIsForfait & "', '" & vartxfac & "','" & valSba & "','" & valDelaiR & "','" & valSommeHT & "', '" & valSommeTTC & "', '" & valLibelle2 & "','" & valDatePoste & "');"
    Debug.Print sql_command
    access_db.Execute (sql_command)
    typ = "FRCT"
    sql_command = "INSERT INTO [FACT] (NUMFACTURE, TYPE, LIBELLE, CLIENT, MONTANTTTC ) VALUES ('" & newnumber & _
    "', '" & typ & "','" & valLibelle & "','" & valClient & "','" & valSommeTTC & "');"
    access_db.Execute (sql_command)
    access_db.Close
    Call rct_layout(valClient, valCollab, valDateF, valDelaiR, valDatePoste, valFonction, valLibelle, valSba, valIsForfait, valSommeHT, valSommeTTC, vartxfac, valLibelle2, newnumber, valisavoir)
End Function
'---------------------------------------------------------------------------------------
' Method : get_last_invoice_num
' Author : p0intz3r0
' Date   : 26/09/2016
' Purpose: Gets the last invoice number, and adds 1 to it
'---------------------------------------------------------------------------------------
Public Function get_last_invoice_num()
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim sql_command As String
    Dim last_invoice_query_result  As Double, last_invoice_query_result2 As Double
    sql_command = "SELECT NUMFACTURE FROM [FACT] ORDER BY NUMFACTURE DESC;"
    Set access_rs = access_db.OpenRecordset(sql_command)
    If Not (access_rs.BOF And access_rs.EOF) Then
        last_invoice_query_result = access_rs.Fields(0)
        Debug.Print last_invoice_query_result
        get_last_invoice_num = last_invoice_query_result + 1
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close

End Function
Public Function get_last_client_num()
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim sql_command As String
    Dim last_client_query_result  As Double, last_invoice_client_result2 As Double
    sql_command = "SELECT REFCLIENT FROM [CLT] ORDER BY REFCLIENT DESC;"
    Set access_rs = access_db.OpenRecordset(sql_command)
    If Not (access_rs.BOF And access_rs.EOF) Then
        last_client_query_result = access_rs.Fields(0)
        Debug.Print last_client_query_result
        get_last_client_num = last_client_query_result + 1
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close

End Function
'---------------------------------------------------------------------------------------
' Method : is_client_factor
' Author : p0intz3r0
' Date   : 26/09/2016
' Purpose: Check if a client should receive a direct invoice
'---------------------------------------------------------------------------------------
Public Function is_client_factor(client)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim is_factor_query_result As Boolean
    Dim sql_command As String
    sql_command = "SELECT FACTORDIRECT FROM [CLT] WHERE CLTNOM = '" & client & "';"
    Set access_rs = access_db.OpenRecordset(sql_command)
    If Not (access_rs.BOF And access_rs.EOF) Then
        is_factor_query_result = access_rs.Fields(0)
        Debug.Print is_factor_query_result
        is_client_factor = is_factor_query_result
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close
    access_db.Close
End Function
'---------------------------------------------------------------------------------------
' Method : admin_logs
' Author : p0intz3r0
' Date   : 26/09/2016
' Purpose: Logs users actions in table [LOG]. windows variable username , timestamp , command and invoice n*
'---------------------------------------------------------------------------------------
Public Function admin_logs(lastcommand, lastnumber)
    Dim access_db As DAO.Database
    Set access_db = DAO.OpenDatabase(db_path)
    Dim is_factor_query_result As Boolean
    Dim sql_command As String
    Dim us3r As String
    us3r = Environ("username")
    sql_command = "INSERT INTO [LOG] (username, timest, command, num) VALUES ('" & us3r & "', '" & Now() & "','" & Left(lastcommand, 8) & "','" & lastnumber & "');"
    Debug.Print sql_command
    access_db.Execute (sql_command)
    admin_logs = "OK"
    access_db.Close
End Function
Public Function get_collab_number(collab)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim sql_command As String
    Dim lastcommand As String, lastnumber As Double
    sql_command = "SELECT MATRICULE FROM [COLLAB] WHERE NOM ='" & collab & "';"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    Dim collab_matricule As Integer
    If Not (access_rs.BOF And access_rs.EOF) Then
        collab_matricule = access_rs.Fields(0)
        Debug.Print collab_matricule
        get_collab_number = collab_matricule
    Else
        get_collab_number = 999
    End If
    access_rs.Close
    access_db.Close
End Function
Public Function get_client_number(client)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim sql_command As String
    Dim lastcommand As String, lastnumber As Double
    sql_command = "SELECT REFCLIENT FROM [CLT] WHERE CLTNOM ='" & client & "';"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    Dim client_matricule As Double
    If Not (access_rs.BOF And access_rs.EOF) Then
        client_matricule = access_rs.Fields(0)
        Debug.Print client_matricule
        get_client_number = client_matricule
    Else
        get_client_number = 0
    End If
    access_rs.Close
    access_db.Close
End Function
Public Function get_collab_name(collab)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim collabnom As String
    Dim sql_command As String
    sql_command = "SELECT NOM, PRENOM FROM [COLLAB] WHERE MATRICULE =" & collab & ";"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    Dim collab_matricule As Integer
    If Not (access_rs.BOF And access_rs.EOF) Then
        collabnom = access_rs.Fields(0) & access_rs.Fields(1)
        Debug.Print collabnom
        get_collab_name = collabnom
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close
    access_db.Close
End Function
Public Function get_client_delai(client)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim sql_command As String
    Dim lastcommand As String, lastnumber As Double
    sql_command = "SELECT DELAI FROM [CLT] WHERE CLTNOM ='" & client & "';"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    Dim client_matricule As Double
    If Not (access_rs.BOF And access_rs.EOF) Then
        client_matricule = access_rs.Fields(0)
        Debug.Print client_matricule
        get_client_delai = client_matricule
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close
    access_db.Close
End Function
Public Function get_client_name(invoice)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim collabnom As String
    Dim sql_command As String
    sql_command = "SELECT CLTNOM FROM [CLT] WHERE CLTNOM =" & invoice.client & ";"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    Dim clientNom As String
    If Not (access_rs.BOF And access_rs.EOF) Then
        clientNom = access_rs.Fields(0)
        Debug.Print collabnom
        get_client_name = clientNom
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close
    access_db.Close
End Function
Public Function get_all_info_RCT(oldinvoice)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Dim sql_command As String
    Dim typeFacture As String
    Set access_db = DAO.OpenDatabase(db_path)
        sql_command = "SELECT RCT_DATEFAC, RCT_CLIENT, RCT_MONTANTHT FROM [FACTRCT] WHERE RCT_ID = " & oldinvoice.number & ";"
        Debug.Print sql_command
        Set access_rs = access_db.OpenRecordset(sql_command)
        If Not (access_rs.BOF And access_rs.EOF) Or access_rs.Fields(0) = Null Then
            oldinvoice.invoicedate = access_rs.Fields(0)
            oldinvoice.client = access_rs.Fields(1)
            oldinvoice.tjm = access_rs.Fields(2)
            oldinvoice.joursfact = 1
        Else
            MsgBox "Erreur. Impossible de retrouver les informations de la facture n°" & oldinvoice.number & vbNewLine & _
                "Contactez votre administrateur"
        End If
       ' oldinvoice.client = get_client_name(oldinvoice)
        Debug.Print sql_command
End Function
Public Function get_client_name2(str_client_name)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim collabnom As String
    Dim sql_command As String
    sql_command = "SELECT CLTNOM FROM [CLT] WHERE REFCLIENT = " & str_client_name & ";"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    Dim clientNom As String
    If Not (access_rs.BOF And access_rs.EOF) Then
        clientNom = access_rs.Fields(0)
        Debug.Print collabnom
        get_client_name2 = clientNom
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close
    access_db.Close
End Function
Public Function get_client_delai_from_num(client)
    Dim access_db As DAO.Database
    Dim access_rs As DAO.Recordset
    Set access_db = DAO.OpenDatabase(db_path)
    Dim sql_command As String
    Dim lastcommand As String, lastnumber As Double
    sql_command = "SELECT DELAI FROM [CLT] WHERE REFCLIENT =" & client & ";"
    Debug.Print sql_command
    Set access_rs = access_db.OpenRecordset(sql_command)
    Dim clt_delai As Double
    If Not (access_rs.BOF And access_rs.EOF) Then
        clt_delai = access_rs.Fields(0)
        Debug.Print clt_delai
        get_client_delai_from_num = clt_delai
    Else
        MsgBox "Erreur "
    End If
    access_rs.Close
    access_db.Close
End Function
