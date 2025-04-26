@echo off
setlocal enabledelayedexpansion

:: Parcourt tous les fichiers DEPT_??.xlsx dans le dossier courant
for %%f in (DEPT_??.xlsx) do (
    set "filename=%%f"
    set "num=!filename:~5,2!"
    ren "%%f" "news_dep_!num!_xxx.xlsx"
)

echo Renommage terminé.
pause
