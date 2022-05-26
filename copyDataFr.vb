Sub copyDataFr()

    ' Déclaration des variables
    
    Dim compteur1 As Long
    Dim dernierLingeFeuil1 As Long, dernierLingeFeuil2 As Long
    Dim mois As String, annee As String

    ' Dans la feuil1 on va copier les données de Rangé A du début
    ' jusqu'à la fin.
    
    dernierLingeFeuil1 = Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row + 1
    dernierLingeFeuil2 = Worksheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    ' Une boucle qui permet de parcourir la feuil1 qui
    ' contient les données, il copie les données si le mois
    ' de Latest Hire Date est egal à 03 donc Mars
     
    For compteur1 = 2 To dernierLingeFeuil1
        
        dernierLingeFeuil2 = Worksheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row + 1
        
        'Recuperation du mois
        mois = Mid(Worksheets("Sheet1").Range("L" & compteur1).Value, 4, 2)
    
        'Recuperation de l année
        annee = Mid(Worksheets("Sheet1").Range("L" & compteur1).Value, 7, 4)
        
        ' Si le mois egal à 03, on stocke dans la feuil2 les données
        ' de la feuil1

        If mois = "03" Then
            
            If annee = "2013" Then
            
                Worksheets("feuil2").Range("A" & dernierLingeFeuil2) = Worksheets("feuil1").Range("A" & compteur1)
                Worksheets("feuil2").Range("B" & dernierLingeFeuil2) = Worksheets("feuil1").Range("B" & compteur1)
                Worksheets("feuil2").Range("C" & dernierLingeFeuil2) = Worksheets("feuil1").Range("C" & compteur1)
                Worksheets("feuil2").Range("D" & dernierLingeFeuil2) = Worksheets("feuil1").Range("D" & compteur1)
                Worksheets("feuil2").Range("E" & dernierLingeFeuil2) = Worksheets("feuil1").Range("E" & compteur1)
                Worksheets("feuil2").Range("F" & dernierLingeFeuil2) = Worksheets("feuil1").Range("F" & compteur1)
                Worksheets("feuil2").Range("G" & dernierLingeFeuil2) = Worksheets("feuil1").Range("G" & compteur1)
                Worksheets("feuil2").Range("H" & dernierLingeFeuil2) = Worksheets("feuil1").Range("H" & compteur1)
                Worksheets("feuil2").Range("I" & dernierLingeFeuil2) = Worksheets("feuil1").Range("I" & compteur1)
                Worksheets("feuil2").Range("J" & dernierLingeFeuil2) = Worksheets("feuil1").Range("J" & compteur1)
                Worksheets("feuil2").Range("K" & dernierLingeFeuil2) = Worksheets("feuil1").Range("K" & compteur1)
                Worksheets("feuil2").Range("L" & dernierLingeFeuil2) = Worksheets("feuil1").Range("L" & compteur1)
                Worksheets("feuil2").Range("M" & dernierLingeFeuil2) = Worksheets("feuil1").Range("M" & compteur1)
                Worksheets("feuil2").Range("N" & dernierLingeFeuil2) = Worksheets("feuil1").Range("N" & compteur1)
                Worksheets("feuil2").Range("O" & dernierLingeFeuil2) = Worksheets("feuil1").Range("O" & compteur1)
                Worksheets("feuil2").Range("P" & dernierLingeFeuil2) = Worksheets("feuil1").Range("P" & compteur1)
            
            End If
          
        End If

    Next compteur1

End Sub
