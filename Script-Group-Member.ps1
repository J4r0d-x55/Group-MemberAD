Import-Module ActiveDirectory

$groupes = Get-ADGroup -Filter * -Properties Members

# Début du chronomètre
$startTime = Get-Date

# Créer un nouvel objet Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Créer un nouveau classeur Excel
$classeur = $excel.Workbooks.Add()
$feuille = $classeur.Worksheets.Item(1)
$feuille.Cells.Item(1, 1) = "Nom du groupe"
$feuille.Cells.Item(1, 2) = "Membre"

# Ligne de départ pour les données dans Excel
$row = 2

# Parcourir chaque groupe
foreach ($groupe in $groupes) {
    # Ajouter le nom du groupe à Excel
    $feuille.Cells.Item($row, 1) = $groupe.Name

    # Vérifier s'il y a des membres dans le groupe
    if ($groupe.Members) {
        foreach ($membre in $groupe.Members) {
            # Récupérer les informations sur le membre
            $infosMembre = Get-ADObject -Identity $membre -Properties Name, SamAccountName
            $feuille.Cells.Item($row, 2) = "$($infosMembre.Name) ($($infosMembre.SamAccountName))"
            $row++
        }
    }
    else {
        $feuille.Cells.Item($row, 2) = "Aucun membre dans ce groupe"
        $row++
    }
}

# Enregistrer le classeur Excel
$cheminFichier = "C:\Users\adm-jbo\Desktop\Group-MemberAD.xlsx"
$classeur.SaveAs($cheminFichier)

# Fermer Excel
$classeur.Close()
$excel.Quit()

# Libérer les objets COM
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($feuille) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($classeur) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Fin du chronomètre
$endTime = Get-Date

# Calcul du temps écoulé
$executionTime = $endTime - $startTime

# Affichage du temps d'exécution et l'emplacement du fichier 
Write-Host "Le script s'est exécuté en $($executionTime.TotalSeconds) secondes."

Write-Host "Export terminé. Le fichier Excel a été enregistré à l'emplacement : $cheminFichier"