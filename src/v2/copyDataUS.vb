Sub copyDataUS()

    ' Déclaration des variables
    
    Dim compteur As Long
    Dim dernierLingeFeuil1 As Long
    Dim dernierLingeFeuil2 As Long
    Dim mois As String
    Dim moisSaisi As String
    Dim annee As String
    Dim anneeSaisi As String
        
    ' On demande de saisir le mois et l'année
    
    moisSaisi = InputBox("Entrer le mois a chercher :", "Saisissez le Mois")
    anneeSaisi = InputBox("Entrer l'annee a chercher :", "Saisissez l'Annee")
    
    ' Dans la feuil1 on va copier les données de Rangé A du début
    ' jusqu'à la fin.
    
    dernierLingeFeuil1 = Worksheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row + 1
    dernierLingeFeuil2 = Worksheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    ' Une boucle qui permet de parcourir la feuil1 qui contient les données
    ' il copie les données si le mois et l'annee qu'on a saisi se trouve
    ' dans la colonne Latest Hire Date
    
    ' On commence à parcourir apartir de la 2e ligne, parce que les données
    ' commence à la ligne numero 2
    
    For compteur = 2 To dernierLingeFeuil1
        
        dernierLingeFeuil2 = Worksheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row + 1
        
        'Recuperation du mois
        mois = Mid(Worksheets("Sheet1").Range("L" & compteur).Value, 6, 2)
        
        'Recuperation de l année
        annee = Mid(Worksheets("Sheet1").Range("L" & compteur).Value, 1, 4)
        
        If annee = anneeSaisi Then
            
            If mois = moisSaisi Then
            
                Worksheets("Sheet2").Range("A" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("A" & compteur)
                Worksheets("Sheet2").Range("B" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("B" & compteur)
                Worksheets("Sheet2").Range("C" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("C" & compteur)
                Worksheets("Sheet2").Range("D" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("D" & compteur)
                Worksheets("Sheet2").Range("E" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("E" & compteur)
                Worksheets("Sheet2").Range("F" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("F" & compteur)
                Worksheets("Sheet2").Range("G" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("G" & compteur)
                Worksheets("Sheet2").Range("H" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("H" & compteur)
                Worksheets("Sheet2").Range("I" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("I" & compteur)
                Worksheets("Sheet2").Range("J" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("J" & compteur)
                Worksheets("Sheet2").Range("K" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("K" & compteur)
                Worksheets("Sheet2").Range("L" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("L" & compteur)
                Worksheets("Sheet2").Range("M" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("M" & compteur)
                Worksheets("Sheet2").Range("N" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("N" & compteur)
                Worksheets("Sheet2").Range("O" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("O" & compteur)
                Worksheets("Sheet2").Range("P" & dernierLingeFeuil2) = Worksheets("Sheet1").Range("P" & compteur)
            
            End If
          
        End If

    Next compteur

End Sub