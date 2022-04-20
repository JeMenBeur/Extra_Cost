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
        Dim dateEcheanceCadence As String
        Dim dateQuaiCadence As String

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

        Dim arrayList = New ArrayList()
        Dim arrayListAllArticlesWithoutCodeFournisseur = New ArrayList()
        Dim arrayListAllArticlesCodeFournisseur = New ArrayList()
        Dim arrayListArticles = New ArrayList()
        Dim arrayListArticlesDetail = New ArrayList()
        Dim arrayListAllArticlesDetailWithoutCodeFournisseur = New ArrayList()
        Dim arrayListAllArticlesDetailCodeFournisseur = New ArrayList()
        Dim arrayListLigneDetail = New ArrayList()

        'Récupère les articles qui sont arrivés en stock au moins 2 fois après le mois de septembre et qui font partis de la liste des 17 codes fournisseur
        Dim strSql = "SELECT POI.ITNBR
                    FROM amflib6.POITEM as POI
                    JOIN amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
                    WHERE POI.DUEDT > 1210901 and POI.STAIC <> 99
                    GROUP BY POI.ITNBR
                    HAVING COUNT(POI.ITNBR) > 1"

        Dim Command As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        Dim RsAs400 = Command.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            arrayListAllArticlesWithoutCodeFournisseur.Add(article)
        Loop

        RsAs400.Close()
        CnnAs400.Close()

        'Récupération des les articles ayant été réception au moins 2 fois depuis septembre et qui font parti de la liste des 17 codes fournisseurs
        strSql = "Select POHI.ITNBR
                    FROM amflib6.POHISTI as POHI
                    JOIN amflib6.POHSTM as POHS on POHS.ORDNO = POHI.ORDNO
                    WHERE POHI.DUEDT > 1210901 and POHI.STAIC <> 99
                    GROUP BY POHI.ITNBR
                    HAVING COUNT(POHI.ITNBR) > 1"
        Dim Command3 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command3.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            arrayListAllArticlesWithoutCodeFournisseur.Add(article)
        Loop

        RsAs400.Close()
        CnnAs400.Close()

        strSql = "SELECT POI.ITNBR
                    FROM amflib6.POITEM as POI
                    JOIN amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
                    WHERE POI.DUEDT > 1210901 and POM.VNDNR in (" & Join(arrayCodeFournisseur.ToArray, ", ") & ") and POI.STAIC <> 99
                    GROUP BY POI.ITNBR
                    HAVING COUNT(POI.ITNBR) > 1"

        Dim Command7 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command7.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            arrayListAllArticlesCodeFournisseur.Add(article)
        Loop

        RsAs400.Close()
        CnnAs400.Close()

        strSql = "SELECT POHI.ITNBR
                    FROM amflib6.POHISTI as POHI
                    JOIN amflib6.POHSTM as POHS on POHS.ORDNO = POHI.ORDNO
                    WHERE POHI.DUEDT > 1210901 and POHS.VNDNR in (" & Join(arrayCodeFournisseur.ToArray, ", ") & ") and POHI.STAIC <> 99
                    GROUP BY POHI.ITNBR
                    HAVING COUNT(POHI.ITNBR) > 1"

        Dim Command8 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command8.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            arrayListAllArticlesCodeFournisseur.Add(article)
        Loop

        RsAs400.Close()
        CnnAs400.Close()

        For Each value In arrayListAllArticlesWithoutCodeFournisseur
            article = value
            For Each valueArticle In arrayListAllArticlesCodeFournisseur
                If valueArticle = article Then
                    arrayListArticles.Add(article)
                End If
            Next
        Next


        'For Each value In arrayListAllArticlesDetailWithoutCodeFournisseur
        '    article = value
        '    For Each valueArticle In arrayListAllArticlesDetailCodeFournisseur
        '        If valueArticle = article Then
        '            arrayListArticlesDetail.Add(article)
        '        End If
        '    Next
        'Next

        ''Récupère les articles qui sont arrivés en stock au moins 2 fois après le mois de septembre et qui font partis de la liste des 17 codes fournisseur
        'strSql = "SELECT POI.ITNBR
        '            FROM amflib6.POITEM as POI
        '            JOIN amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
        '            WHERE POI.DUEDT > 1210901 and POM.VNDNR in (" & Join(arrayCodeFournisseur.ToArray, ", ") & ") and POI.STAIC <> 99
        '                and POI.ITNBR in (" & Join(arrayListAllArticles.ToArray, ", ") & ")"

        'Dim Command6 As New OdbcCommand(strSql, CnnAs400)
        'CnnAs400.Open()
        'RsAs400 = Command6.ExecuteReader()

        'Do While RsAs400.Read()
        '    article = RsAs400.GetValue(0).ToString
        '    arrayListArticles.Add(article)
        'Loop

        'RsAs400.Close()
        'CnnAs400.Close()

        'Récupère les colonnes des articles qui sont arrivés en stock au moins 2 fois après le mois de septembre et qui font partis de la liste des 17 codes fournisseur
        strSql = "SELECT POI.ITNBR, POM.ORDNO, POM.ACTDT, POM.CURID, V.VNAME, POI.ITNBR, POI.QTYOR, POI.BLCOD, POI.QTDEV, 
                        POI.ACTPR, POI.ACTPL, POI.STAIC, POI.DUEDT, POI.DOKDT, POM.VNDNR
                    FROM amflib6.POITEM as POI
                    JOIN amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
                    JOIN amflib6.VENNAM as V on V.VNDNR = POM.VNDNR
                    WHERE POI.DUEDT > 1210901 and POI.STAIC <> 99 and POI.ITNBR in ('" & Join(arrayListArticles.ToArray, "', '") & "')"
        Dim Command2 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command2.ExecuteReader()

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
            codeFournisseur = RsAs400.GetValue(14).ToString

            Dim arrayLigne = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCommandeOrigine, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
                                    situation, dateStock, dateQuai, codeFournisseur}
            arrayList.Add(arrayLigne)
        Loop

        RsAs400.Close()
        CnnAs400.Close()



        'Récupération les colonnes des articles ayant été réception au moins 2 fois depuis septembre et qui font parti de la liste des 17 codes fournisseurs
        strSql = "Select POHI.ITNBR, POHS.ORDNO, POHS.ACTDT, POHS.CURID, V.VNAME, POHI.ITNBR, POHI.QTYOR, POHI.BLCOD, POHI.QTDEV, POHI.ACTPR, POHI.ACTPL, POHI.STAIC, POHI.DUEDT, 
                    POHI.DOKDT, POHS.VNDNR 
                    FROM amflib6.POHISTI as POHI
                    JOIN amflib6.POHSTM as POHS on POHS.ORDNO = POHI.ORDNO
                    JOIN amflib6.VENNAM as V on V.VNDNR = POHS.VNDNR
                    WHERE POHI.DUEDT > 1210901 and POHI.STAIC <> 99 and POHI.ITNBR in ('" & Join(arrayListArticles.ToArray, "', '") & "')
                    ORDER BY POHI.ITNBR"
        Dim Command4 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command4.ExecuteReader()

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
            codeFournisseur = RsAs400.GetValue(14).ToString

            Dim arrayLigne = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCommandeOrigine, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
                                    situation, dateStock, dateQuai, codeFournisseur}
            arrayList.Add(arrayLigne)
        Loop

        RsAs400.Close()
        CnnAs400.Close()

        For Each value In arrayList
            'If (value(4).ToString.Trim = "THA92097000") Then
            '    MsgBox("ok")
            'End If
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
