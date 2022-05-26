Sub copierDonnerUs()

    ' Déclaration des variables
    
    Dim compteur1 As Long
    Dim dernierLingeFeuil1 As Long, dernierLingeFeuil2 As Long
    Dim mois As String

    ' Dans la feuil1 on va copier les données de Rangé A du début
    ' jusqu'à la fin.
    
    dernierLingeFeuil1 = Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row + 1
    dernierLingeFeuil2 = Worksheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    ' Une boucle qui permet de parcourir la feuil1 qui
    ' contient les données, il copie les données si le mois
    ' de Latest Hire Date est egal à 3 donc Mars
     
    For compteur1 = 2 To dernierLingeFeuil1
        
        dernierLingeFeuil2 = Worksheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row + 1
        mois = Mid(Worksheets("Sheet1").Range("L" & compteur1).Value, 1, 1)
        
        ' Si le mois egal à 3, on stocke dans la feuil2 les données
        ' de la feuil1

        If mois = "3" Then
            Worksheets("Sheet2").Range("A" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("A" & compteur1)
            Worksheets("Sheet2").Range("B" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("B" & compteur1)
            Worksheets("Sheet2").Range("C" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("C" & compteur1)
            Worksheets("Sheet2").Range("D" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("D" & compteur1)
            Worksheets("Sheet2").Range("E" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("E" & compteur1)
            Worksheets("Sheet2").Range("F" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("F" & compteur1)
            Worksheets("Sheet2").Range("G" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("G" & compteur1)
            Worksheets("Sheet2").Range("H" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("H" & compteur1)
            Worksheets("Sheet2").Range("I" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("I" & compteur1)
            Worksheets("Sheet2").Range("J" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("J" & compteur1)
            Worksheets("Sheet2").Range("K" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("K" & compteur1)
            Worksheets("Sheet2").Range("L" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("L" & compteur1)
            Worksheets("Sheet2").Range("M" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("M" & compteur1)
            Worksheets("Sheet2").Range("N" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("N" & compteur1)
            Worksheets("Sheet2").Range("O" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("O" & compteur1)
            Worksheets("Sheet2").Range("P" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("P" & compteur1)
        End If

    Next compteur1

End Sub



