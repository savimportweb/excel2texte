# excel2texte.exe

Utilisation : 
        -f"Chemin Du fichier XLS/XLSX"
        -s"Nom de la Feuille"
PAS ENCORE IMPLEMENTE :-dNuméro de la feuille
        -xy"Position X;Y de départ", example : A1={0;0}
t -o output dans le meme dossier avec meme nom mais .csv
        -v Version
        
Excel2texte.exe -f"c:\temp.xls" -o



Si pas de nom de feuille choisi, il prends la premiere. si pas -o, l'output est redirigé dans output.texte

il faut juste les 3.dll et l'executable tourne avec le framework 3.5. pas testé sous tous mais globalement ca marche.
