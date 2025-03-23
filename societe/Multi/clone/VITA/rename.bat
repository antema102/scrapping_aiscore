@echo off
setlocal enabledelayedexpansion

:: Boucle sur tous les fichiers correspondant au motif
for %%F in (news_dep_*_xxx.xlsx) do (
    set "ancienNom=%%F"
    
    :: Extraire le numéro entre 'news_dep_' et le premier '_'
    for /f "tokens=3 delims=_." %%A in ("%%F") do (
        set "numero=%%A"
        set "nouveauNom=DEPT_!numero!.xlsx"

        :: Renommer le fichier
        ren "!ancienNom!" "!nouveauNom!"
        echo Renommé: "!ancienNom!" → "!nouveauNom!"
    )
)

echo.
echo ✅ Tous les fichiers ont été renommés avec succès !
pause
