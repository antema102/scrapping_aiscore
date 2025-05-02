@echo off
setlocal enabledelayedexpansion

rem Boucle de 1 à 99
for /L %%i in (1,1,99) do (
    rem Formater le nombre avec un zéro devant si inférieur à 10
    set "num=%%i"
    if %%i LSS 10 set "num=0%%i"

    rem Renommer le fichier s'il existe
    if exist "dep_!num!_.xlsx" (
        ren "dep_!num!_.xlsx" "news_dep_!num!.xlsx"
        echo Renommé : dep_!num!_.xlsx → news_dep_!num!.xlsx
    ) else (
        echo Fichier non trouvé : dep_!num!_.xlsx
    )
)

pause
✅ Instructions