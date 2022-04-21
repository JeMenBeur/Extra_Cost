<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.numeroCommande = New System.Windows.Forms.ColumnHeader()
        Me.poste = New System.Windows.Forms.ColumnHeader()
        Me.nomFournisseur = New System.Windows.Forms.ColumnHeader()
        Me.dateCommande = New System.Windows.Forms.ColumnHeader()
        Me.devise = New System.Windows.Forms.ColumnHeader()
        Me.article = New System.Windows.Forms.ColumnHeader()
        Me.quantiteCommandeOrigine = New System.Windows.Forms.ColumnHeader()
        Me.quantiteDeviation = New System.Windows.Forms.ColumnHeader()
        Me.prixUnitaireEuro = New System.Windows.Forms.ColumnHeader()
        Me.PrixUnitaireDevise = New System.Windows.Forms.ColumnHeader()
        Me.situation = New System.Windows.Forms.ColumnHeader()
        Me.dateStock = New System.Windows.Forms.ColumnHeader()
        Me.dateQuai = New System.Windows.Forms.ColumnHeader()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.numeroCommande, Me.poste, Me.nomFournisseur, Me.dateCommande, Me.devise, Me.article, Me.quantiteCommandeOrigine, Me.quantiteDeviation, Me.prixUnitaireEuro, Me.PrixUnitaireDevise, Me.situation, Me.dateStock, Me.dateQuai})
        Me.ListView1.FullRowSelect = True
        Me.ListView1.Location = New System.Drawing.Point(12, 62)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(1298, 485)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'numeroCommande
        '
        Me.numeroCommande.Text = "commande"
        Me.numeroCommande.Width = 100
        '
        'poste
        '
        Me.poste.Text = "poste"
        '
        'nomFournisseur
        '
        Me.nomFournisseur.Text = "fournisseur"
        Me.nomFournisseur.Width = 220
        '
        'dateCommande
        '
        Me.dateCommande.Text = "date commande"
        Me.dateCommande.Width = 100
        '
        'devise
        '
        Me.devise.Text = "devise"
        '
        'article
        '
        Me.article.Text = "article"
        Me.article.Width = 110
        '
        'quantiteCommandeOrigine
        '
        Me.quantiteCommandeOrigine.Text = "Qte commande origine"
        Me.quantiteCommandeOrigine.Width = 140
        '
        'quantiteDeviation
        '
        Me.quantiteDeviation.Text = "Qte deviation"
        Me.quantiteDeviation.Width = 90
        '
        'prixUnitaireEuro
        '
        Me.prixUnitaireEuro.Text = "PU Euro"
        Me.prixUnitaireEuro.Width = 90
        '
        'PrixUnitaireDevise
        '
        Me.PrixUnitaireDevise.Text = "PU devise"
        Me.PrixUnitaireDevise.Width = 90
        '
        'situation
        '
        Me.situation.Text = "situation"
        '
        'dateStock
        '
        Me.dateStock.Text = "date stock"
        Me.dateStock.Width = 80
        '
        'dateQuai
        '
        Me.dateQuai.Text = "date quai"
        Me.dateQuai.Width = 80
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1322, 609)
        Me.Controls.Add(Me.ListView1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListView1 As ListView
    Friend WithEvents numeroCommande As ColumnHeader
    Friend WithEvents nomFournisseur As ColumnHeader
    Friend WithEvents dateCommande As ColumnHeader
    Friend WithEvents devise As ColumnHeader
    Friend WithEvents article As ColumnHeader
    Friend WithEvents quantiteCommandeOrigine As ColumnHeader
    Friend WithEvents quantiteDeviation As ColumnHeader
    Friend WithEvents prixUnitaireEuro As ColumnHeader
    Friend WithEvents PrixUnitaireDevise As ColumnHeader
    Friend WithEvents situation As ColumnHeader
    Friend WithEvents dateStock As ColumnHeader
    Friend WithEvents dateQuai As ColumnHeader
    Friend WithEvents poste As ColumnHeader
End Class
