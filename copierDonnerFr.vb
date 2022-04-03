Sub copierDonnerFr()

    ' Déclaration des variables

    Dim compteur1 As Long
    Dim dernierLingeFeuil1 As Long, dernierLingeFeuil2 As Long
    Dim mois As String

    ' Dans la feuil1 on va copier les données de Rangé A du début
    ' jusqu'à la fin.
    
    dernierLingeFeuil1 = Worksheets("feuil1").Range("A" & Rows.Count).End(xlUp).Row + 1
    dernierLingeFeuil2 = Worksheets("feuil2").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    ' Une boucle qui permet de parcourir la feuil1 qui
    ' contient les données, il copie les données si le mois
    ' de la cellule Date est egal à 03 donc Mars

    For compteur1 = 2 To dernierLingeFeuil1
    
        dernierLingeFeuil2 = Worksheets("feuil2").Range("A" & Rows.Count).End(xlUp).Row + 1
        mois = Mid(Worksheets("Feuil1").Range("L" & compteur1).Value, 4, 2)
        
        ' Si le mois egal à 03, on stocke dans la feuil2 les données
        ' de la feuil1
        
        If mois = "03" Then
            
        End If
        
    Next compteur1

End Sub