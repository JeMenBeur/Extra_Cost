Imports System.Data.Odbc
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    Private sheetXls As Worksheet

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim article, numeroCommande, codeFournisseur, nomFournisseur, devise, dateCommande As String
        Dim quantiteCommandeOrigine, quantiteDeviation As Double
        Dim prixUnitaireEuro, prixUnitaireDevise As Double
        Dim situation, dateStock, dateQuai As String
        Dim quantiteCadence As Double
        Dim dateQuaiCadence, dateStockCadence As String

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
        Dim CnnAs400 = New Odbc.OdbcConnection(chconnect)

        Dim arrayList = New ArrayList()
        Dim arrayListArticles = New ArrayList()
        Dim arrayListArticlesDetail = New ArrayList()

        'Récupération des articles dont la date stock est supérieur au 1er septembre
        Dim strSql = "SELECT POI.ITNBR, POM.VNDNR
                    FROM amflib6.POITEM as POI
                    JOIN amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
                    WHERE POI.DUEDT > 1210901 and POI.STAIC <> 99"

        Dim Command As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        Dim RsAs400 = Command.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            codeFournisseur = RsAs400.GetValue(1).ToString
            arrayListArticles.Add({article, codeFournisseur})
        Loop

        RsAs400.Close()
        CnnAs400.Close()


        'Récupération des articles historiques dont la date stock est supérieur au 1er septembre
        strSql = "SELECT POHI.ITNBR, POHS.VNDNR
                FROM amflib6.POHISTI as POHI
                JOIN amflib6.POHSTM as POHS on POHS.ORDNO = POHI.ORDNO
                WHERE POHI.DUEDT > 1210901 and POHI.STAIC <> 99"

        Dim Command1 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command1.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            codeFournisseur = RsAs400.GetValue(1).ToString
            arrayListArticles.Add({article, codeFournisseur})
        Loop

        RsAs400.Close()
        CnnAs400.Close()

        'Récupération des cadences dont la date stock est supérieur au 1er septembre
        strSql = "SELECT POB.ITNBR, POM.VNDNR
                FROM amflib6.POBLKT as POB
                JOIN amflib6.POMAST as POM on POM.ORDNO = POB.ORDNO
                WHERE POB.RELDT > 1210901 and POB.STAIC <> 99"

        Dim Command2 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command2.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            codeFournisseur = RsAs400.GetValue(1).ToString
            arrayListArticles.Add({article, codeFournisseur})
        Loop

        RsAs400.Close()
        CnnAs400.Close()


        'Récupération des cadences historiques dont la date stock est supérieur au 1er septembre
        strSql = "SELECT POHI.ITNBR, POHS.VNDNR
                FROM amflib6.POHISTB as POHI
                JOIN amflib6.POHSTM as POHS on POHS.ORDNO = POHI.ORDNO
                WHERE POHI.RELDT > 1210901 and POHI.STAIC <> 99"

        Dim Command3 As New OdbcCommand(strSql, CnnAs400)
        CnnAs400.Open()
        RsAs400 = Command3.ExecuteReader()

        Do While RsAs400.Read()
            article = RsAs400.GetValue(0).ToString
            codeFournisseur = RsAs400.GetValue(1).ToString
            arrayListArticles.Add({article, codeFournisseur})
        Loop

        RsAs400.Close()
        CnnAs400.Close()

        'Parmi l'ensemble des articles, sauvegarde ceux dont le code fournisseur fait parti des 17 codes du tableau 
        For Each value In arrayListArticles
            article = value(0)
            codeFournisseur = value(1)
            For Each fournisseur In arrayCodeFournisseur
                If fournisseur = Trim(codeFournisseur) Then
                    arrayList.Add(article)
                End If
            Next
        Next

        'Supprime les doublons du tableau et enregistre les valeurs dans un nouveau tableau
        For j = 0 To arrayList.Count - 1
            article = arrayList(j)
            If article <> "" Then
                arrayListArticlesDetail.Add(article)
            End If
            For x = 0 To arrayList.Count - 1
                If article = arrayList(x) And j <> x Then
                    arrayList(x) = ""
                End If
            Next
        Next

        'Boucle sur tous les articles du tableau
        For Each value In arrayListArticlesDetail
            'Récupère les colonnes des articles qui sont arrivés en stock au moins 1 fois après le mois de septembre et qui font partis de la liste des 17 codes fournisseur
            strSql = "SELECT POI.ITNBR, POM.ORDNO, POM.ACTDT, POM.CURID, V.VNAME, POI.QTYOR, POI.QTDEV, 
                        POI.EXTPR, POI.EXTPL, POI.STAIC, POI.DUEDT, POI.DOKDT, POM.VNDNR
                    FROM amflib6.POITEM as POI
                    JOIN amflib6.POMAST as POM on POM.ORDNO = POI.ORDNO
                    JOIN amflib6.VENNAM as V on V.VNDNR = POM.VNDNR
                    WHERE POI.DUEDT > 1210901 and POI.STAIC <> 99 and POI.ITNBR = '" & value & "'
                    ORDER BY POI.ITNBR"
            Dim Command4 As New OdbcCommand(strSql, CnnAs400)
            CnnAs400.Open()
            RsAs400 = Command4.ExecuteReader()

            Do While RsAs400.Read()
                article = RsAs400.GetValue(0).ToString
                numeroCommande = RsAs400.GetValue(1).ToString
                dateCommande = RsAs400.GetValue(2).ToString
                devise = RsAs400.GetValue(3).ToString
                nomFournisseur = RsAs400.GetValue(4).ToString
                quantiteCommandeOrigine = If(RsAs400.GetValue(5).ToString <> "", RsAs400.GetValue(5).ToString, 0)
                quantiteDeviation = If(RsAs400.GetValue(6).ToString <> "", RsAs400.GetValue(6).ToString, 0)
                prixUnitaireEuro = If(RsAs400.GetValue(7).ToString <> "", RsAs400.GetValue(7).ToString, 0)
                prixUnitaireDevise = If(RsAs400.GetValue(8).ToString <> "", RsAs400.GetValue(8).ToString, 0)
                situation = RsAs400.GetValue(9).ToString
                dateStock = RsAs400.GetValue(10).ToString
                dateQuai = RsAs400.GetValue(11).ToString
                codeFournisseur = RsAs400.GetValue(12).ToString

                Dim arrayLigne = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCommandeOrigine, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
                                        situation, dateStock, dateQuai, codeFournisseur, ""}
                arrayList.Add(arrayLigne)
            Loop

            RsAs400.Close()
            CnnAs400.Close()

            'Récupération les colonnes des articles historique ayant été réception au moins 1 fois depuis septembre et qui font parti de la liste des 17 codes fournisseurs
            strSql = "Select POHI.ITNBR, POHS.ORDNO, POHS.ACTDT, POHS.CURID, V.VNAME, POHI.QTYOR, POHI.QTDEV, POHI.EXTPR, POHI.EXTPL, POHI.STAIC, POHI.DUEDT, 
                        POHI.DOKDT, POHS.VNDNR 
                    FROM amflib6.POHISTI as POHI
                    JOIN amflib6.POHSTM as POHS on POHS.ORDNO = POHI.ORDNO
                    JOIN amflib6.VENNAM as V on V.VNDNR = POHS.VNDNR
                    WHERE POHI.DUEDT > 1210901 and POHI.STAIC <> 99 and POHI.ITNBR = '" & value & "'
                    ORDER BY POHI.ITNBR"
            Dim Command5 As New OdbcCommand(strSql, CnnAs400)
            CnnAs400.Open()
            RsAs400 = Command5.ExecuteReader()

            Do While RsAs400.Read()
                article = RsAs400.GetValue(0).ToString
                numeroCommande = RsAs400.GetValue(1).ToString
                dateCommande = RsAs400.GetValue(2).ToString
                devise = RsAs400.GetValue(3).ToString
                nomFournisseur = RsAs400.GetValue(4).ToString
                quantiteCommandeOrigine = If(RsAs400.GetValue(5).ToString <> "", RsAs400.GetValue(5).ToString, 0)
                quantiteDeviation = If(RsAs400.GetValue(6).ToString <> "", RsAs400.GetValue(6).ToString, 0)
                prixUnitaireEuro = If(RsAs400.GetValue(7).ToString <> "", RsAs400.GetValue(7).ToString, 0)
                prixUnitaireDevise = If(RsAs400.GetValue(8).ToString <> "", RsAs400.GetValue(8).ToString, 0)
                situation = RsAs400.GetValue(9).ToString
                dateStock = RsAs400.GetValue(10).ToString
                dateQuai = RsAs400.GetValue(11).ToString
                codeFournisseur = RsAs400.GetValue(12).ToString

                Dim arrayLigne = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCommandeOrigine, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
                                        situation, dateStock, dateQuai, codeFournisseur, ""}
                arrayList.Add(arrayLigne)
            Loop

            RsAs400.Close()
            CnnAs400.Close()


            'Récupère les colonnes des cadences des articles qui sont arrivés en stock au moins 1 fois après le mois de septembre et qui font partis de la liste des 17 codes fournisseur
            strSql = "SELECT POB.ITNBR, POB.ORDNO, POB.RELQT, POB.RELDT, POB.STAIC, POB.DOKDT, POB.EXTPR, POB.EXTPL, V.VNDNR, V.VNAME, POM.CURID, POM.ACTDT
                    FROM amflib6.POBLKT as POB
                    JOIN amflib6.POMAST as POM on POM.ORDNO = POB.ORDNO
                    JOIN amflib6.VENNAM as V on V.VNDNR = POM.VNDNR
                    WHERE POB.DOKDT > 1210901 and POB.STAIC <> 99 and POB.ITNBR = '" & value & "'
                    ORDER BY POB.ITNBR"
            Dim Command6 As New OdbcCommand(strSql, CnnAs400)
            CnnAs400.Open()
            RsAs400 = Command6.ExecuteReader()

            Do While RsAs400.Read()
                article = RsAs400.GetValue(0).ToString
                numeroCommande = RsAs400.GetValue(1).ToString
                quantiteCadence = If(RsAs400.GetValue(2).ToString <> "", RsAs400.GetValue(2).ToString, 0)
                dateStockCadence = RsAs400.GetValue(3).ToString
                situation = RsAs400.GetValue(4).ToString
                dateQuaiCadence = RsAs400.GetValue(5).ToString
                prixUnitaireEuro = If(RsAs400.GetValue(6).ToString <> "", RsAs400.GetValue(6).ToString, 0)
                prixUnitaireDevise = If(RsAs400.GetValue(7).ToString <> "", RsAs400.GetValue(7).ToString, 0)
                codeFournisseur = RsAs400.GetValue(8).ToString
                nomFournisseur = RsAs400.GetValue(9).ToString
                devise = RsAs400.GetValue(10).ToString
                dateCommande = RsAs400.GetValue(11).ToString
                quantiteDeviation = 0
                Dim arrayCadence = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCadence, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
                                situation, dateStockCadence, dateQuaiCadence, codeFournisseur, 0}
                arrayList.Add(arrayCadence)
            Loop

            RsAs400.Close()
            CnnAs400.Close()


            'Récupère les colonnes des cadences historiques qui sont arrivés en stock au moins 1 fois après le mois de septembre et qui font partis de la liste des 17 codes fournisseur
            strSql = "SELECT POHI.ITNBR, POHI.ORDNO, POHI.RELQT, POHI.RELDT, POHI.STAIC, POHI.DOKDT, POHI.EXTPR, POHI.EXTPL, V.VNDNR, V.VNAME, POHS.CURID, POHS.ACTDT
                    FROM amflib6.POHISTB as POHI
                    JOIN amflib6.POHSTM as POHS on POHS.ORDNO = POHI.ORDNO
                    JOIN amflib6.VENNAM as V on V.VNDNR = POHS.VNDNR
                    WHERE POHI.DOKDT > 1210901 and POHI.STAIC <> 99 and POHI.ITNBR = '" & value & "'
                    ORDER BY POHI.ITNBR"
            Dim Command7 As New OdbcCommand(strSql, CnnAs400)
            CnnAs400.Open()
            RsAs400 = Command7.ExecuteReader()

            Do While RsAs400.Read()
                article = RsAs400.GetValue(0).ToString
                numeroCommande = RsAs400.GetValue(1).ToString
                quantiteCadence = If(RsAs400.GetValue(2).ToString <> "", RsAs400.GetValue(2).ToString, 0)
                dateStockCadence = RsAs400.GetValue(3).ToString
                situation = RsAs400.GetValue(4).ToString
                dateQuaiCadence = RsAs400.GetValue(5).ToString
                prixUnitaireEuro = If(RsAs400.GetValue(6).ToString <> "", RsAs400.GetValue(6).ToString, 0)
                prixUnitaireDevise = If(RsAs400.GetValue(7).ToString <> "", RsAs400.GetValue(7).ToString, 0)
                codeFournisseur = RsAs400.GetValue(8).ToString
                nomFournisseur = RsAs400.GetValue(9).ToString
                devise = RsAs400.GetValue(10).ToString
                dateCommande = RsAs400.GetValue(11).ToString
                quantiteDeviation = 0
                Dim arrayCadence = {numeroCommande, nomFournisseur, dateCommande, devise, article, quantiteCadence, quantiteDeviation, prixUnitaireEuro, prixUnitaireDevise,
                                    situation, dateStockCadence, dateQuaiCadence, codeFournisseur, 1}
                arrayList.Add(arrayCadence)
            Loop

            RsAs400.Close()
            CnnAs400.Close()
        Next

        Dim xls As Application
        Dim xlsfeuille As Worksheet
        Dim xlsclasseur As Workbook
        Dim pchar As String

        xls = CreateObject("Excel.Application")
        xlsclasseur = xls.Workbooks.Add
        xlsfeuille = xlsclasseur.Worksheets(1)

        Dim i = 2
        For Each value In arrayList
            xlsfeuille.Cells(i, 1) = value(0)
            xlsfeuille.Cells(i, 2) = value(13)
            xlsfeuille.Cells(i, 3) = value(1)
            xlsfeuille.Cells(i, 4) = value(2)
            xlsfeuille.Cells(i, 5) = value(3)
            xlsfeuille.Cells(i, 6) = value(4)
            xlsfeuille.Cells(i, 7) = value(5)
            xlsfeuille.Cells(i, 8) = value(6)
            xlsfeuille.Cells(i, 9) = value(7)
            xlsfeuille.Cells(i, 10) = value(8)
            xlsfeuille.Cells(i, 11) = value(9)
            xlsfeuille.Cells(i, 12) = value(10)
            xlsfeuille.Cells(i, 13) = value(11)
            i += 1
            If i = 100 Then
                MsgBox(i)
            End If
            If i = 500 Then
                MsgBox(i)
            End If
            If i = 1000 Then
                MsgBox(i)
            End If
            If i = 2000 Then
                MsgBox(i)
            End If
            If i = 3000 Then
                MsgBox(i)
            End If
            If i = 4000 Then
                MsgBox(i)
            End If
        Next

        xlsclasseur.SaveAs("d:\Classeur2.xlsx")
        xls.Application.Quit()
        xls = Nothing
    End Sub

End Class