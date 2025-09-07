param(
    [string]$original_docx,
    [string]$revised_docx,
    [string]$comments_csv,
    [string]$output_docx
)

# Vérification des paramètres
if (-not ($original_docx -and $revised_docx -and $comments_csv -and $output_docx)) {
    Write-Error "Usage: compare_and_comment.ps1 -original_docx <path> -revised_docx <path> -comments_csv <path> -output_docx <path>"
    exit 1
}

# Créer l'objet Word COM
try {
    $word = New-Object -ComObject Word.Application
}
catch {
    Write-Error "Microsoft Word n'est pas installé ou une erreur est survenue lors de son lancement."
    exit 1
}

$word.Visible = $false

try {
    # Ouvrir le document révisé
    $doc = $word.Documents.Open($revised_docx)

    # Activer le suivi des modifications
    $doc.TrackRevisions = $true

    # Comparer avec le document original
    $doc.Compare($original_docx)

    # Mettre en évidence en vert toutes les insertions issues de la comparaison
    $wdRevisionInsert = 1        # WdRevisionType.wdRevisionInsert
    $wdBrightGreen    = 4        # WdColorIndex.wdBrightGreen
    foreach ($rev in $doc.Revisions) {
        try {
            if ($rev.Type -eq $wdRevisionInsert) {
                $rev.Range.HighlightColorIndex = $wdBrightGreen
            }
        } catch {
            # Ignorer les erreurs de mise en forme ponctuelles
        }
    }

    # Injecter les commentaires depuis le CSV
    $comments = Import-Csv -Path $comments_csv -Encoding UTF8
    foreach ($comment in $comments) {
        $ancre = $comment.ancre_textuelle
        $texte_commentaire = $comment.commentaire
        $gravite = $comment.gravite
        $categorie = $comment.categorie

        $range = $doc.Content
        $find = $range.Find
        $find.Text = $ancre
        $find.Forward = $true
        $find.Wrap = 1 # wdFindContinue
        $find.MatchCase = $false
        $find.MatchWholeWord = $false
        $find.MatchWildcards = $false
        $find.MatchSoundsLike = $false
        $find.MatchAllWordForms = $false

        if ($find.Execute()) {
            $comment_text = "[$gravite - $categorie] $texte_commentaire"
            $doc.Comments.Add($range, $comment_text)
        }
    }

    # Enregistrer le document final
    $doc.SaveAs([ref]$output_docx)
    $doc.Close()

    Write-Host "Le document final a été généré avec succès: $output_docx"
}
catch {
    Write-Error "Une erreur est survenue lors du traitement des documents Word: $_"
}
finally {
    # Quitter Word
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
