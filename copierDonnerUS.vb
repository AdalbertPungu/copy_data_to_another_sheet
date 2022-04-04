Sub copierDonnerUs()

    ' Déclaration des variables
    
    Dim compteur1 As Long
    Dim dernierLingeFeuil1 As Long, dernierLingeFeuil2 As Long
    Dim mois As String

    ' Dans la feuil1 on va copier les données de Rangé A du début
    ' jusqu'à la fin.
    
    dernierLingeFeuil1 = Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row + 1
    dernierLingeFeuil2 = Worksheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    
End Sub



