Imports System.Data.Odbc

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim article As String
        Dim numeroCommande As String
        Dim codeFournisseur As String
        Dim nomFournisseur As String
        Dim devise As String
        Dim dateCommande As String
        Dim quantiteCommandeOrigine As Double
        Dim quantiteDeviation As Double
        Dim blanketItemCode As String
        Dim prixUnitaireEuro As Double
        Dim prixUnitaireDevise As Double
        Dim situation As String
        Dim dateStock As String
        Dim dateQuai As String
        Dim quantiteCadence As Double

        Dim arrayCodeFournisseur = New List(Of String)
        arrayCodeFournisseur.Add("500265")
        arrayCodeFournisseur.Add("2716")
        arrayCodeFournisseur.Add("500072")
        arrayCodeFournisseur.Add("500346")
        arrayCodeFournisseur.Add("500064")
        arrayCodeFournisseur.Add("500996")
        arrayCodeFournisseur.Add("118039")
        arrayCodeFournisseur.Add("501419")
        arrayCodeFournisseur.Add("501467")
        arrayCodeFournisseur.Add("500424")
        arrayCodeFournisseur.Add("500606")
        arrayCodeFournisseur.Add("501431")
        arrayCodeFournisseur.Add("501706")
        arrayCodeFournisseur.Add("501088")
        arrayCodeFournisseur.Add("501291")
        arrayCodeFournisseur.Add("500961")
        arrayCodeFournisseur.Add("501476")

        Dim chconnect As String = "Dsn=APMXAC;uid=APMXAC;pwd=MXAC"
        Dim CnnAs400 = New OdbcConnection(chconnect)

        ''Récupère le numéro de commande, codefournisseur, date commande, devise et nom fournisseur des commandes faites après le premier septembre et dont le code fournisseur fait parti du tableau
        'Dim strSql = "Select POM.ORDNO, POM.VNDNR, POM.ACTDT, POM.CURID, V.VNAME from amflib6.POMAST as POM
        '            Join amflib6.VENNAM as V on V.VNDNR = POM.VNDNR
        '            Join amflib6.POITEM as POI on POI.ORDNO = POM.ORDNO
        '            Where POI.DUEDT > 1210901 and POM.VNDNR in (" & Join(arrayCodeFournisseur.ToArray, ", ") & ")"
        'Dim Command As New OdbcCommand(strSql, CnnAs400)
        'CnnAs400.Open()
        'Dim RsAs400 As OdbcDataReader = Command.ExecuteReader()
        'Dim arrayListCommande = New ArrayList()
        'Do While RsAs400.Read()
        '    numeroCommande = RsAs400.GetValue(0).ToString
        '    codeFournisseur = RsAs400.GetValue(1).ToString
        '    dateCommande = RsAs400.GetValue(2).ToString
        '    devise = RsAs400.GetValue(3).ToString
        '    nomFournisseur = RsAs400.GetValue(4).ToString

        '    Dim arrayCommande = {numeroCommande, codeFournisseur, nomFournisseur, dateCommande, devise}
        '    arrayListCommande.Add(arrayCommande)
        'Loop

        'RsAs400.Close()
        'CnnAs400.Close()

        Dim arrayList = New ArrayList()
        Dim arrayListLigneDetail = New ArrayList()

        'Récupère toutes les lignes des commandes récupéré dans la précédente requête
        'For Each value In arrayListCommande
        Dim strSql = "Select POI.ITNBR, POM.ORDNO, POM.ACTDT, POM.CURID, V.VNAME, POI.ITNBR, POI.QTYOR, POI.BLCOD, POI.QTDEV, POI.ACTPR, POI.ACTPL, POI.STAIC, POI.DUEDT, 
                    POI.DOKDT from amflib6.POITEM as POI
                    join amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
                    join amflib6.VENNAM as V on V.VNDNR = POM.VNDNR
                    where POI.DUEDT > 1210901 and POM.VNDNR in (" & Join(arrayCodeFournisseur.ToArray, ", ") & ")"
        Dim Command2 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        Dim RsAs400 = Command2.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            numeroCommande = RsAs400.GetValue(1).ToString
            dateCommande = RsAs400.GetValue(2).ToString
            devise = RsAs400.GetValue(3).ToString
            nomFournisseur = RsAs400.GetValue(4).ToString
            article = RsAs400.GetValue(5).ToString
            quantiteCommandeOrigine = If(RsAs400.GetValue(6).ToString <> "", RsAs400.GetValue(6).ToString, 0)
            blanketItemCode = RsAs400.GetValue(7).ToString
            quantiteDeviation = If(RsAs400.GetValue(8).ToString <> "", RsAs400.GetValue(8).ToString, 0)
            prixUnitaireEuro = If(RsAs400.GetValue(9).ToString <> "", RsAs400.GetValue(9).ToString, 0)
            prixUnitaireDevise = If(RsAs400.GetValue(10).ToString <> "", RsAs400.GetValue(10).ToString, 0)
            situation = RsAs400.GetValue(11).ToString
            dateStock = RsAs400.GetValue(12).ToString
            dateQuai = RsAs400.GetValue(13).ToString

            Dim arrayLigne = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCommandeOrigine, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
                                    situation, dateStock, dateQuai}
            'Dim arrayLigne = {value(0), value(1), value(4), value(2), value(3), article, quantiteCommandeOrigine, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
            '                    situation, dateStock, dateQuai}
            arrayList.Add(arrayLigne)
        Loop

        RsAs400.Close()
        CnnAs400.Close()
        'Next

        ''Récupère toutes les informations des articles qui ont été commandés par 2 fournisseur différents parmis la liste des codes fournisseurs du tableau
        'For Each value In arrayList
        '    numeroCommande = value(0)
        '    codeFournisseur = value(1)
        '    nomFournisseur = value(2)
        '    dateCommande = value(3)
        '    devise = value(4)
        '    article = value(5)
        '    strSql = "Select POI.ITNBR, POI.QTYOR, POI.BLCOD, POI.QTDEV, POI.ACTPR, POI.ACTPL, POI.STAIC, POI.DUEDT, 
        '            POI.DOKDT from amflib6.POITEM as POI
        '            join amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
        '            where POI.ITNBR = '" & article & "' and POI.DUEDT > 1210901 and POM.VNDNR <> " & codeFournisseur & " and POM.VNDNR in (" & Join(arrayCodeFournisseur.ToArray, ", ") & ")"
        '    Dim Command3 As New OdbcCommand(strSql, CnnAs400)
        '    CnnAs400.Open()
        '    RsAs400 = Command3.ExecuteReader()

        '    Do While RsAs400.Read()
        '        article = RsAs400.GetValue(0).ToString
        '        quantiteCommandeOrigine = If(RsAs400.GetValue(1).ToString <> "", RsAs400.GetValue(1).ToString, 0)
        '        blanketItemCode = RsAs400.GetValue(2).ToString
        '        quantiteDeviation = If(RsAs400.GetValue(3).ToString <> "", RsAs400.GetValue(3).ToString, 0)
        '        prixUnitaireEuro = If(RsAs400.GetValue(4).ToString <> "", RsAs400.GetValue(4).ToString, 0)
        '        prixUnitaireDevise = If(RsAs400.GetValue(5).ToString <> "", RsAs400.GetValue(5).ToString, 0)
        '        situation = RsAs400.GetValue(6).ToString
        '        dateStock = RsAs400.GetValue(7).ToString
        '        dateQuai = RsAs400.GetValue(8).ToString

        '        Dim arrayLigne = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCommandeOrigine, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
        '                            situation, dateStock, dateQuai}
        '        arrayListLigneDetail.Add(arrayLigne)
        '    Loop

        '    RsAs400.Close()
        '    CnnAs400.Close()
        'Next

        For Each value In arrayList
            If (value(4).ToString.Trim = "THA92162248") Then
                MsgBox("OK")
            End If
            Dim _MyListViewItem As ListViewItem = ListView1.Items.Add(value(0))
            With _MyListViewItem
                .SubItems.Add(value(1))
                .SubItems.Add(value(2))
                .SubItems.Add(value(3))
                .SubItems.Add(value(4))
                .SubItems.Add(value(5))
                .SubItems.Add(value(6))
                .SubItems.Add(value(7))
                .SubItems.Add(value(8))
                .SubItems.Add(value(9))
                .SubItems.Add(value(10))
                .SubItems.Add(value(11))
            End With
        Next

    End Sub

End Class
