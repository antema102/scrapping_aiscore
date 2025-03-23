@echo off
setlocal enabledelayedexpansion

:: Boucle sur tous les fichiers de DEP_01.xlsx à DEP_92.xlsx
for %%F in (news_dep_??.xlsx) do (
    set "ancienNom=%%F"
    
    :: Extraire le numéro après 'DEP_' et avant '.xlsx'
    for /f "tokens=3 delims=_." %%A in ("%%F") do (
        set "numero=%%A"
        set "nouveauNom=dep_!numero!_.xlsx"

        :: Renommer le fichier
        ren "!ancienNom!" "!nouveauNom!"
        echo Renommé: "!ancienNom!" → "!nouveauNom!"
    )
)

echo.
echo ✅ Tous les fichiers ont été renommés avec succès !
pause
