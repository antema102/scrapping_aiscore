@echo off
setlocal enabledelayedexpansion

REM Boucle pour renommer les fichiers de dep_01_sources.xlsx à DEPT_01.xlsx
for /l %%i in (1,1,99) do (
    REM Ajout du zéro devant les numéros inférieurs à 10
    set num=%%i
    if %%i lss 10 set num=0%%i

    REM Création du nom de fichier original et du nouveau nom
    set oldFile=dep_!num!_sources.xlsx
    set newFile=DEPT_!num!.xlsx

    REM Vérifier si le fichier existe, puis le renommer
    if exist "!oldFile!" (
        ren "!oldFile!" "!newFile!"
        echo Renommé !oldFile! en !newFile!
    )
)

echo Fin du script.
pause
